"""
Travel Accommodation Portal (internal network)
Flask + pyodbc (SQL Server) + MSAL/Microsoft Graph mail + signed approval links
"""
import html
import logging
import os
import re
import secrets
import uuid
from datetime import date, datetime

import msal
import pyodbc
import requests
from dotenv import load_dotenv
from flask import Flask, Response, abort, jsonify, render_template, request
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

load_dotenv()

# --------------------------------------------------------------------------
# Configuration (all from environment / App Settings)
# --------------------------------------------------------------------------
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

# Mailbox the app sends from (the app registration needs Graph Mail.Send for it)
SENDER_MAILBOX = os.environ.get("SENDER_MAILBOX", "")

# Intranet base URL managers open from the approval email, no trailing slash, e.g. http://10.10.30.90:8080
APP_BASE_URL = os.environ.get("APP_BASE_URL", "").rstrip("/")

# Waitress bind address. 0.0.0.0 = reachable from the internal network on every interface of this server
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8080"))

# Signs the Approve/Reject links in manager emails. Must stay constant, or links already sent stop working.
SECRET_KEY = os.environ.get("SECRET_KEY", "")
REVIEW_LINK_MAX_AGE = int(os.environ.get("REVIEW_LINK_MAX_AGE_DAYS", "30")) * 86400

# Travel desk (Jigna, Mittal, Sagar) and CC (Pooja) - comma separated email addresses
TRAVEL_DESK_TO = [e.strip() for e in os.environ.get("TRAVEL_DESK_TO", "").split(",") if e.strip()]
TRAVEL_DESK_CC = [e.strip() for e in os.environ.get("TRAVEL_DESK_CC", "").split(",") if e.strip()]

# SQL Server
DB_DRIVER = os.environ.get("DB_DRIVER", "ODBC Driver 18 for SQL Server")
DB_SERVER = os.environ.get("DB_SERVER", "")
DB_NAME = os.environ.get("DB_NAME", "")
DB_USER = os.environ.get("DB_USER", "")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "")
DB_PORT = os.environ.get("DB_PORT", "1433")
DB_TRUST_CERT = os.environ.get("DB_TRUST_SERVER_CERT", "yes")  # "yes" for self-signed certs
# Optional: a full ODBC connection string overrides the individual DB_* values above
DB_CONNECTION_STRING = os.environ.get("DB_CONNECTION_STRING", "")

# Hierarchy table (EmpCode -> ManagerCode)
MAPPING_TABLE = os.environ.get("MAPPING_TABLE", "dbo.checklist_empmanager_mapping")
# Email columns inside the mapping table itself
EMP_EMAIL_COL = os.environ.get("EMP_EMAIL_COL", "EmpEmail")
MGR_EMAIL_COL = os.environ.get("MGR_EMAIL_COL", "ManagerEmail")

# HR master (name, email, status, Department/Designation ids) and the id -> name lookup tables
EMPLOYEE_TABLE = os.environ.get("EMPLOYEE_TABLE", "dbo.tblEmployeeMaster")
DEPARTMENT_TABLE = os.environ.get("DEPARTMENT_TABLE", "dbo.checklists_department")
DESIGNATION_TABLE = os.environ.get("DESIGNATION_TABLE", "dbo.checklists_designation")

# Optional login gate for the form (HTTP Basic). Leave PORTAL_PASSWORD empty to disable (local dev only).
PORTAL_USER = os.environ.get("PORTAL_USER", "")
PORTAL_PASSWORD = os.environ.get("PORTAL_PASSWORD", "")

# Optional image for the header. Empty (default) shows the built-in Hunger Pangs mark.
LOGO_URL = os.environ.get("LOGO_URL", "")

ROOM_CATEGORIES = {"Single Occupancy", "Double Occupancy"}
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("travel-portal")

if not SECRET_KEY:
    SECRET_KEY = secrets.token_urlsafe(32)
    log.warning("SECRET_KEY is not set: approval links will stop working when the server restarts")
_review_signer = URLSafeTimedSerializer(SECRET_KEY, salt="travel-review")


# --------------------------------------------------------------------------
# Database helpers
# --------------------------------------------------------------------------
_IDENT = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)?$")


def _q(identifier):
    """Validate and bracket-quote a table/column name that comes from configuration."""
    if not _IDENT.match(identifier):
        raise RuntimeError(f"Invalid SQL identifier in configuration: {identifier!r}")
    return ".".join(f"[{part}]" for part in identifier.split("."))


def get_conn():
    if DB_CONNECTION_STRING:
        return pyodbc.connect(DB_CONNECTION_STRING, timeout=15)
    conn_str = (
        f"DRIVER={{{DB_DRIVER}}};SERVER={DB_SERVER},{DB_PORT};DATABASE={DB_NAME};"
        f"UID={DB_USER};PWD={DB_PASSWORD};"
        f"Encrypt=yes;TrustServerCertificate={DB_TRUST_CERT};Connection Timeout=15;"
    )
    return pyodbc.connect(conn_str)


def _clean(v):
    return str(v).strip() if v is not None and str(v).strip() else None


def fetch_employee(emp_code):
    """
    Employee details from the HR master plus their reporting manager from the mapping table.
    Emails come from the mapping table, falling back to the HR master.
    Returns None if the code is in neither table. NeedsApproval is False when the employee
    has no reporting manager on record (HODs / senior staff).
    """
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT e.EmpCode, e.EmployeeName, e.EmployeeEmail, e.Status, dep.name, des.name,
                   m.ManagerCode, m.EmpEmail, m.ManagerEmail,
                   mgr.EmployeeName, mgr.EmployeeEmail, mdes.name
            FROM (SELECT CAST(? AS nvarchar(30)) AS EmpCode) k
            LEFT JOIN {_q(EMPLOYEE_TABLE)} e ON e.EmpCode = k.EmpCode
            LEFT JOIN {_q(DEPARTMENT_TABLE)} dep ON CAST(dep.id AS varchar(20)) = e.Department
            LEFT JOIN {_q(DESIGNATION_TABLE)} des ON CAST(des.id AS varchar(20)) = e.Designation
            OUTER APPLY (
                SELECT TOP 1 ManagerCode, {_q(EMP_EMAIL_COL)} AS EmpEmail, {_q(MGR_EMAIL_COL)} AS ManagerEmail
                FROM {_q(MAPPING_TABLE)} WHERE EmpCode = k.EmpCode ORDER BY id DESC
            ) m
            LEFT JOIN {_q(EMPLOYEE_TABLE)} mgr ON mgr.EmpCode = m.ManagerCode
            LEFT JOIN {_q(DESIGNATION_TABLE)} mdes ON CAST(mdes.id AS varchar(20)) = mgr.Designation
            """,
            emp_code,
        )
        row = cur.fetchone()
    (master_code, name, email, status, department, designation,
     mgr_code, map_emp_email, map_mgr_email, mgr_name, mgr_email, mgr_designation) = row
    mgr_code = _clean(mgr_code)
    if not master_code and not mgr_code:
        return None
    return {
        "EmpCode": _clean(master_code) or emp_code,  # HR casing, e.g. "h189" -> "H189"
        "EmpName": _clean(name),
        "EmpEmail": _clean(map_emp_email) or _clean(email),
        "Active": not master_code or _clean(status) == "Active",
        "Designation": _clean(designation),
        "Department": _clean(department),
        "NeedsApproval": bool(mgr_code),
        "ManagerCode": mgr_code,
        "ManagerName": _clean(mgr_name) or mgr_code,
        "ManagerDesignation": _clean(mgr_designation),
        "ManagerEmail": _clean(map_mgr_email) or _clean(mgr_email),
    }


def employee_problem(emp):
    """A reason this employee cannot submit yet (shown at lookup time), or None."""
    if not emp["Active"]:
        return "This employee code is not active in HR records."
    if emp["NeedsApproval"] and not emp["ManagerEmail"]:
        return (f"Your manager ({emp['ManagerName']}) has no email address on record, so the approval "
                "email cannot be sent. Please ask HR to update it.")
    if not emp["EmpEmail"]:
        return "Your email address is not on record. Please ask HR to update it."
    return None


def insert_request(req):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO TravelRequests
              (RequestId, EmpCode, EmpName, EmpEmail, ManagerCode, ManagerName, ManagerEmail,
               Designation, Department, FromDate, ToDate, PlaceFrom, PlaceTo,
               RoomCategory, PreferredLocation, ReasonOfTravel, Status)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'Pending')
            """,
            req["RequestId"], req["EmpCode"], req["EmpName"], req["EmpEmail"],
            req["ManagerCode"], req["ManagerName"], req["ManagerEmail"],
            req["Designation"], req["Department"], req["FromDate"], req["ToDate"],
            req["PlaceFrom"], req["PlaceTo"], req["RoomCategory"], req["PreferredLocation"],
            req["ReasonOfTravel"],
        )
        conn.commit()


def set_status(request_id, status):
    with get_conn() as conn:
        conn.cursor().execute(
            "UPDATE TravelRequests SET Status = ? WHERE RequestId = ?", status, request_id
        )
        conn.commit()


def load_request(request_id):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM TravelRequests WHERE RequestId = ?", request_id)
        row = cur.fetchone()
        if not row:
            return None
        cols = [c[0] for c in cur.description]
        return dict(zip(cols, row))


def claim_decision(request_id, status, comment):
    """Atomically move Pending -> Approved/Rejected. Returns True if this call won."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE TravelRequests SET Status = ?, ActionedAt = SYSUTCDATETIME(), ActionComment = ?, "
            "RejectionReason = ? WHERE RequestId = ? AND Status = 'Pending'",
            status, comment, (comment if status == "Rejected" else None), request_id,
        )
        won = cur.rowcount == 1
        conn.commit()
        return won


# --------------------------------------------------------------------------
# Microsoft Graph helpers
# --------------------------------------------------------------------------
_msal_app = None


def get_graph_token():
    global _msal_app
    if _msal_app is None:
        _msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )
    result = _msal_app.acquire_token_for_client(scopes=GRAPH_SCOPE)  # MSAL caches tokens
    if "access_token" not in result:
        raise RuntimeError(f"Graph token error: {result.get('error_description', result)}")
    return result["access_token"]


def send_mail(to, subject, html_body, cc=None):
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
            "ccRecipients": [{"emailAddress": {"address": a}} for a in (cc or [])],
        },
        "saveToSentItems": True,
    }
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{SENDER_MAILBOX}/sendMail",
        headers={"Authorization": f"Bearer {get_graph_token()}", "Content-Type": "application/json"},
        json=payload,
        timeout=30,
    )
    if resp.status_code != 202:
        raise RuntimeError(f"sendMail failed {resp.status_code}: {resp.text}")


# --------------------------------------------------------------------------
# Presentation helpers
# --------------------------------------------------------------------------
def fmt_date(d):
    if isinstance(d, (date, datetime)):
        return d.strftime("%d-%b-%Y")
    return str(d)


def details_table_html(r):
    rows = [
        ("Emp Code", r["EmpCode"]),
        ("Name", r["EmpName"]),
        ("Designation", r["Designation"]),
        ("Department", r["Department"]),
        ("From (Date)", fmt_date(r["FromDate"])),
        ("To (Date)", fmt_date(r["ToDate"])),
        ("Travelling From", r["PlaceFrom"]),
        ("Travelling To", r["PlaceTo"]),
        ("Reason of Travel", r.get("ReasonOfTravel") or "-"),
        ("Category of Room", r["RoomCategory"]),
        ("Preferred Location", r["PreferredLocation"] or "-"),
    ]
    body = "".join(
        f"<tr><td style='padding:4px 12px 4px 0;font-weight:600'>{html.escape(k)}</td>"
        f"<td style='padding:4px 0'>{html.escape(str(v))}</td></tr>"
        for k, v in rows
    )
    return f"<table style='font-family:Segoe UI,Arial,sans-serif;font-size:14px'>{body}</table>"


# --------------------------------------------------------------------------
# Manager approval links (intranet)
# --------------------------------------------------------------------------
def make_review_token(request_id):
    return _review_signer.dumps(str(request_id))


def check_review_token(request_id, token):
    """True if the token was issued for this request and has not expired."""
    try:
        return _review_signer.loads(token, max_age=REVIEW_LINK_MAX_AGE) == str(request_id)
    except (SignatureExpired, BadSignature):
        return False


def review_url(request_id, decision):
    return f"{APP_BASE_URL}/review/{request_id}?t={make_review_token(request_id)}&decision={decision}"


def _email_button(label, url, color):
    return (
        f"<a href='{html.escape(url)}' style='display:inline-block;padding:10px 26px;margin-right:10px;"
        f"background:{color};color:#ffffff;text-decoration:none;font-weight:600;border-radius:6px;"
        f"font-family:Segoe UI,Arial,sans-serif;font-size:14px'>{html.escape(label)}</a>"
    )


def approval_email_html(r):
    """Manager email: request details plus Approve / Reject links to the intranet review page."""
    rid = r["RequestId"]
    return (
        "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px'>"
        f"<p>{html.escape(r['EmpName'])} ({html.escape(r['EmpCode'])}) has requested travel and "
        "accommodation and needs your approval.</p>"
        + details_table_html(r)
        + "<p style='margin:20px 0'>"
        + _email_button("Approve", review_url(rid, "approve"), "#16a34a")
        + _email_button("Reject", review_url(rid, "reject"), "#dc2626")
        + "</p><p style='color:#64748b;font-size:12px'>The buttons open the Travel Portal on the office "
        "network, where you confirm your decision. A reason is required to reject. "
        "They work only from the office network or VPN.</p></div>"
    )


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
# /review/ links carry their own signed token; /healthz is for monitoring; /static/ is the logo
_OPEN_PREFIXES = ("/review/", "/healthz", "/static/")


@app.before_request
def require_portal_login():
    if not PORTAL_PASSWORD or request.path.startswith(_OPEN_PREFIXES):
        return None
    auth = request.authorization
    if (auth
            and secrets.compare_digest((auth.username or "").encode(), PORTAL_USER.encode())
            and secrets.compare_digest((auth.password or "").encode(), PORTAL_PASSWORD.encode())):
        return None
    return Response("Login required", 401, {"WWW-Authenticate": 'Basic realm="Travel Portal"'})


@app.get("/")
def index():
    today = date.today().isoformat()
    return render_template("index.html", today=today, logo_url=LOGO_URL)


@app.get("/healthz")
def healthz():
    return "ok", 200


@app.get("/api/employee/<emp_code>")
def api_employee(emp_code):
    if not re.fullmatch(r"[A-Za-z0-9\-_/.]{1,30}", emp_code):
        return jsonify({"found": False, "error": "Invalid employee code"}), 400
    try:
        emp = fetch_employee(emp_code.strip())
    except Exception:  # noqa: BLE001
        log.exception("Employee lookup failed")
        return jsonify({"found": False, "error": "Lookup failed"}), 500
    if not emp:
        return jsonify({"found": False}), 404
    return jsonify({
        "found": True,
        "empName": emp["EmpName"] or "",
        "designation": emp["Designation"] or "",
        "department": emp["Department"] or "",
        "needsApproval": emp["NeedsApproval"],
        "managerName": emp["ManagerName"] if emp["NeedsApproval"] else None,
        "managerDesignation": emp["ManagerDesignation"] if emp["NeedsApproval"] else None,
        "problem": employee_problem(emp),
    })


@app.post("/submit-travel")
def submit_travel():
    f = request.form if request.form else (request.get_json(silent=True) or {})
    get = lambda k: str(f.get(k, "")).strip()  # noqa: E731

    emp_code = get("emp_code")
    designation, department = get("designation"), get("department")
    place_from, place_to = get("place_from"), get("place_to")
    room, pref_loc = get("room_category"), get("preferred_location")
    travel_reason = get("travel_reason")

    missing = [n for n, v in {
        "Emp Code": emp_code, "Designation": designation, "Department": department,
        "From (Date)": get("from_date"), "To (Date)": get("to_date"),
        "Travelling From": place_from, "Travelling To": place_to, "Reason of Travel": travel_reason,
        "Category of Room": room,
    }.items() if not v]
    if missing:
        return jsonify({"ok": False, "error": "Missing: " + ", ".join(missing)}), 400

    try:
        from_date = datetime.strptime(get("from_date"), "%Y-%m-%d").date()
        to_date = datetime.strptime(get("to_date"), "%Y-%m-%d").date()
    except ValueError:
        return jsonify({"ok": False, "error": "Invalid date format"}), 400
    if to_date < from_date:
        return jsonify({"ok": False, "error": "'To' date cannot be before 'From' date"}), 400
    if room not in ROOM_CATEGORIES:
        return jsonify({"ok": False, "error": "Invalid room category"}), 400
    if len(travel_reason) > 500:
        return jsonify({"ok": False, "error": "Reason of Travel is too long (max 500 characters)"}), 400
    if max(len(designation), len(department), len(place_from), len(place_to), len(pref_loc)) > 150:
        return jsonify({"ok": False, "error": "A field is too long (max 150 characters)"}), 400

    try:
        emp = fetch_employee(emp_code)
    except Exception:  # noqa: BLE001
        log.exception("DB error during employee fetch")
        return jsonify({"ok": False, "error": "Database error. Please try again."}), 500
    if not emp:
        return jsonify({"ok": False, "error": f"Employee code '{emp_code}' not found"}), 404
    problem = employee_problem(emp)
    if problem:
        return jsonify({"ok": False, "error": problem}), 422
    emp["EmpName"] = emp["EmpName"] or get("emp_name")
    if not emp["EmpName"]:
        return jsonify({"ok": False, "error": "Name of Person travelling is required"}), 400

    # HR-record values (name, designation, department, manager) win over what the browser sent
    req = {
        "RequestId": str(uuid.uuid4()),
        **{k: emp[k] for k in ("EmpCode", "EmpName", "EmpEmail", "ManagerCode")},
        "ManagerName": emp["ManagerName"] if emp["NeedsApproval"] else None,
        "ManagerEmail": emp["ManagerEmail"] if emp["NeedsApproval"] else "",  # column is NOT NULL
        "Designation": emp["Designation"] or designation, "Department": emp["Department"] or department,
        "FromDate": from_date, "ToDate": to_date,
        "PlaceFrom": place_from, "PlaceTo": place_to,
        "RoomCategory": room, "PreferredLocation": pref_loc,
        "ReasonOfTravel": travel_reason,
    }

    try:
        insert_request(req)
    except Exception:  # noqa: BLE001
        log.exception("DB error during insert")
        return jsonify({"ok": False, "error": "Could not save the request."}), 500

    if not emp["NeedsApproval"]:
        # No reporting manager on record (HOD / senior staff): straight to the Travel Desk
        claim_decision(req["RequestId"], "Approved", DIRECT_NOTE)
        try:
            notify_decision(load_request(req["RequestId"]), "Approved", "")
        except Exception:  # noqa: BLE001
            log.exception("Failed to email the Travel Desk")
            set_status(req["RequestId"], "MailFailed")
            return jsonify({"ok": False, "error": "Could not email the Travel Desk. Please contact support."}), 502
        return jsonify({
            "ok": True, "requestId": req["RequestId"], "needsApproval": False,
            "message": "Request submitted and sent directly to the Travel Desk.",
        })

    try:
        send_mail(
            [req["ManagerEmail"]],
            f"Approval required: Travel request for {req['EmpName']} ({req['EmpCode']})",
            approval_email_html(req),
        )
    except Exception:  # noqa: BLE001
        log.exception("Failed to send approval email")
        set_status(req["RequestId"], "MailFailed")
        return jsonify({"ok": False, "error": "Could not email your manager. Please contact support."}), 502

    return jsonify({
        "ok": True,
        "requestId": req["RequestId"],
        "needsApproval": True,
        "message": f"Request submitted. Approval email sent to {req['ManagerName']}.",
    })


DIRECT_NOTE = "No reporting manager on record: sent directly to the Travel Desk"


def notify_decision(r, new_status, comment):
    """Email the Travel Desk (on approval) and the employee after a decision.
    Requests from employees with no reporting manager (ManagerCode empty) skip manager approval."""
    direct = not r.get("ManagerCode")
    note_label = "Reason for rejection" if new_status == "Rejected" else "Manager comment"
    notes = f"<p><b>{note_label}:</b> {html.escape(comment)}</p>" if comment else ""
    if new_status == "Approved":
        how = ("submitted by an employee with <b>no reporting manager on record</b> (HOD / senior staff), "
               "so no manager approval was required" if direct
               else f"<b>approved</b> by {html.escape(r['ManagerName'])}")
        send_mail(
            TRAVEL_DESK_TO,
            f"Travel booking required: {r['EmpName']} ({r['EmpCode']}) - {r['PlaceFrom']} to {r['PlaceTo']}",
            "<p style='font-family:Segoe UI,Arial,sans-serif'>Dear Jigna, Mittal and Sagar,</p>"
            f"<p>The following travel request was {how}. "
            "Please arrange ticketing and accommodation.</p>"
            + details_table_html(r) + notes,
            cc=TRAVEL_DESK_CC,
        )
        emp_msg = (f"<p>Dear {html.escape(r['EmpName'])},</p><p>Your travel request has been "
                   + ("sent to the Travel Desk.</p>" if direct
                      else "<b>approved</b> and forwarded to the Travel Desk.</p>"))
        subject = "Travel request sent to the Travel Desk" if direct else "Travel request approved"
    else:
        emp_msg = (f"<p>Dear {html.escape(r['EmpName'])},</p><p>Your travel request has been "
                   f"<b>rejected</b> by {html.escape(r['ManagerName'])}.</p>")
        subject = "Travel request rejected"
    send_mail([r["EmpEmail"]], subject,
              "<div style='font-family:Segoe UI,Arial,sans-serif'>" + emp_msg
              + details_table_html(r) + notes + "</div>")


@app.route("/review/<request_id>", methods=["GET", "POST"])
def review(request_id):
    """Manager approval page. GET only shows the request (safe for mail link scanners); POST decides."""
    token = request.values.get("t", "")
    try:
        uuid.UUID(request_id)
    except ValueError:
        abort(404)
    if not check_review_token(request_id, token):
        return render_template("review.html", state="invalid", logo_url=LOGO_URL), 403

    r = load_request(request_id)  # trust the DB record only
    if not r:
        abort(404)
    page = dict(r=r, token=token, fmt_date=fmt_date, logo_url=LOGO_URL)

    if request.method == "GET":
        if r["Status"] != "Pending":
            return render_template("review.html", state="already", **page)
        decision = request.args.get("decision", "")
        return render_template("review.html", state="pending", decision=decision, comment="", **page)

    decision = request.form.get("decision", "").lower()
    comment = request.form.get("comment", "").strip()[:300]
    if decision not in ("approve", "reject"):
        return render_template("review.html", state="pending", decision="", comment=comment,
                               error="Choose Approve or Reject.", **page), 400
    if decision == "reject" and not comment:
        return render_template("review.html", state="pending", decision="reject", comment="",
                               error="A reason is required to reject this request.", **page), 400

    new_status = "Approved" if decision == "approve" else "Rejected"
    if not claim_decision(request_id, new_status, comment):
        page["r"] = load_request(request_id)
        return render_template("review.html", state="already", **page)

    try:
        notify_decision(r, new_status, comment)
    except Exception:  # noqa: BLE001
        log.exception("Post-decision email failed for %s", request_id)
    page["r"] = load_request(request_id)
    return render_template("review.html", state="done", **page)


if __name__ == "__main__":
    from waitress import serve
    print(f"Serving Travel Accommodation Portal on http://{HOST}:{PORT} (links use {APP_BASE_URL or 'APP_BASE_URL not set'})",
          flush=True)
    serve(app, host=HOST, port=PORT, threads=8)
