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
from decimal import Decimal, InvalidOperation

import msal
import pyodbc
import requests
from dotenv import load_dotenv
from flask import Flask, Response, abort, jsonify, redirect, render_template, request, url_for
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

# Workflow: employee submits -> Travel Desk books -> manager approves/rejects
STATUS_BOOKING = "Pending Booking"            # waiting for the Travel Desk to arrange the booking
STATUS_MANAGER = "Pending Manager Approval"   # booking arranged, waiting for the manager
DIRECT_NOTE = "No reporting manager on record: approved when the Travel Desk confirmed the booking"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("travel-portal")

if not SECRET_KEY:
    SECRET_KEY = secrets.token_urlsafe(32)
    log.warning("SECRET_KEY is not set: approval links will stop working when the server restarts")
# Separate salts: a Travel Desk link can never be used as a manager approval link, and vice versa
_review_signer = URLSafeTimedSerializer(SECRET_KEY, salt="travel-review")
_desk_signer = URLSafeTimedSerializer(SECRET_KEY, salt="travel-desk")


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
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            req["RequestId"], req["EmpCode"], req["EmpName"], req["EmpEmail"],
            req["ManagerCode"], req["ManagerName"], req["ManagerEmail"],
            req["Designation"], req["Department"], req["FromDate"], req["ToDate"],
            req["PlaceFrom"], req["PlaceTo"], req["RoomCategory"], req["PreferredLocation"],
            req["ReasonOfTravel"], STATUS_BOOKING,
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


def advance_status(request_id, from_status, to_status):
    """Atomically move a request from one status to the next. Returns True if this call won."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("UPDATE TravelRequests SET Status = ? WHERE RequestId = ? AND Status = ?",
                    to_status, request_id, from_status)
        won = cur.rowcount == 1
        conn.commit()
        return won


def save_booking(request_id, area, property_name, price, to_status, comment=None):
    """Store the Travel Desk's booking details and move Pending Booking -> to_status in one atomic update.
    to_status is Pending Manager Approval, or Approved when there is no reporting manager.
    Returns True if this call won."""
    approved = to_status == "Approved"
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE TravelRequests SET BookingArea = ?, PropertyName = ?, Cost = ?, Status = ?"
            + (", ActionedAt = SYSUTCDATETIME(), ActionComment = ?" if approved else "")
            + " WHERE RequestId = ? AND Status = ?",
            area, property_name, price, to_status, *([comment] if approved else []), request_id, STATUS_BOOKING,
        )
        won = cur.rowcount == 1
        conn.commit()
        return won


def claim_decision(request_id, status, comment, from_status=None):
    """Atomically move a request awaiting a decision to Approved/Rejected. Returns True if this call won."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE TravelRequests SET Status = ?, ActionedAt = SYSUTCDATETIME(), ActionComment = ?, "
            "RejectionReason = ? WHERE RequestId = ? AND Status = ?",
            status, comment, (comment if status == "Rejected" else None), request_id,
            from_status or STATUS_MANAGER,
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


def booking_ref(r):
    """Short reference for emails and pages, e.g. '#10001'. Falls back to the GUID prefix before migration 003."""
    if r.get("BookingRefID"):
        return f"#{r['BookingRefID']}"
    return "#" + str(r["RequestId"])[:8].upper()


def with_ref(r, subject):
    return f"Booking Ref: {booking_ref(r)} - {subject}"


# Email building blocks. Inline styles only (email clients ignore <style> blocks); tables are 100% wide up to
# 600px so they fill a phone screen and stay readable on desktop.
_TABLE = ("role='presentation' width='100%' cellpadding='0' cellspacing='0' "
          "style='width:100%;max-width:600px;border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif'")
_BOX = ("role='presentation' width='100%' cellpadding='0' cellspacing='0' "
        "style='width:100%;max-width:600px;border-collapse:separate;margin:0 0 18px;border-radius:12px;{}'")


def ref_badge_html(r):
    return ("<p style='margin:0 0 16px'><span style='display:inline-block;padding:5px 12px;border-radius:999px;"
            "background:#f7efdd;color:#681313;font-weight:700;font-size:13px'>"
            f"Booking Ref: {html.escape(booking_ref(r))}</span></p>")


def details_table_html(r, booking=True):
    """Employee and trip details. booking=False leaves out the Travel Desk's booking (shown in its own box)."""
    rows = [
        ("Booking Ref", booking_ref(r)),
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
    if booking and r.get("PropertyName"):
        rows += booking_rows(r)
    body = "".join(
        "<tr>"
        "<td width='40%' style='width:40%;padding:9px 12px 9px 0;border-bottom:1px solid #efe7d5;color:#7a6f66;"
        f"font-size:13px;vertical-align:top'>{html.escape(k)}</td>"
        "<td style='padding:9px 0;border-bottom:1px solid #efe7d5;color:#1d1d1d;font-size:14px;font-weight:600;"
        f"vertical-align:top;word-break:break-word'>{html.escape(str(v))}</td>"
        "</tr>"
        for k, v in rows
    )
    return f"<table {_TABLE}>{body}</table>"


def booking_box_html(r, title="Booking arranged by the Travel Desk"):
    """The Travel Desk's booking (Place/Area, Property Name, Cost) in a soft, brand-tinted box."""
    rows = "".join(
        "<tr>"
        f"<td width='40%' style='width:40%;padding:6px 12px 6px 0;color:#8a5a4c;font-size:13px'>{html.escape(k)}</td>"
        "<td style='padding:6px 0;color:#1d1d1d;font-size:15px;font-weight:700;word-break:break-word'>"
        f"{html.escape(v)}</td></tr>"
        for k, v in booking_rows(r)
    )
    return (
        f"<table {_BOX.format('background:#fdf3ec;border:1px solid #f3dccf')}>"
        "<tr><td style='padding:14px 18px;font-family:Segoe UI,Arial,sans-serif'>"
        "<div style='font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#c54b36;"
        f"margin-bottom:6px'>{html.escape(title)}</div>"
        f"<table {_TABLE}>{rows}</table>"
        "</td></tr></table>"
    )


def reason_box_html(label, text):
    """The manager's reason/comment, highlighted."""
    return (
        f"<table {_BOX.format('background:#fbeae6;border:1px solid #f0c4b9')}>"
        "<tr><td style='padding:14px 18px;font-family:Segoe UI,Arial,sans-serif'>"
        "<div style='font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#9b3523;"
        f"margin-bottom:4px'>{html.escape(label)}</div>"
        f"<div style='font-size:15px;color:#1d1d1d;word-break:break-word'>{html.escape(text)}</div>"
        "</td></tr></table>"
    )


app.jinja_env.globals["booking_ref"] = booking_ref  # templates show the same reference


def email_html(*parts):
    """Wrap an email body: 600px max, full width on phones."""
    return f"<div style='{EMAIL_FONT};max-width:600px;line-height:1.5'>" + "".join(parts) + "</div>"


def fmt_inr(amount):
    """Indian digit grouping: 123456.5 -> '₹1,23,456.50'."""
    whole, frac = f"{Decimal(amount):.2f}".split(".")
    head, tail = whole[:-3], whole[-3:]
    while len(head) > 2:
        tail = head[-2:] + "," + tail
        head = head[:-2]
    return "₹" + (head + "," + tail if head else tail) + "." + frac


def booking_rows(r):
    """The Travel Desk's booking details as (label, value) pairs."""
    return [
        ("Place/Area", r.get("BookingArea") or "-"),
        ("Property Name", r.get("PropertyName") or "-"),
        ("Cost", fmt_inr(r["Cost"]) if r.get("Cost") is not None else "-"),
    ]


# --------------------------------------------------------------------------
# Signed action links (intranet) and workflow emails - plain HTML, no Actionable Messages
# --------------------------------------------------------------------------
def make_token(signer, request_id):
    return signer.dumps(str(request_id))


def check_token(signer, request_id, token):
    """True if the token was issued by this signer for this request and has not expired."""
    try:
        return signer.loads(token, max_age=REVIEW_LINK_MAX_AGE) == str(request_id)
    except (SignatureExpired, BadSignature):
        return False


def desk_url(request_id):
    return f"{APP_BASE_URL}/travel-desk-confirm/{request_id}?t={make_token(_desk_signer, request_id)}"


def review_url(request_id, decision):
    return f"{APP_BASE_URL}/review/{request_id}?t={make_token(_review_signer, request_id)}&decision={decision}"


EMAIL_FONT = "font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1d1d1d"
LINK_NOTE = ("<p style='color:#7a6f66;font-size:12px'>The button opens the Travel Portal on the office "
             "network, where you confirm. It works only from the office network or VPN.</p>")


def _email_button(label, url, color):
    return (
        f"<a href='{html.escape(url)}' style='display:inline-block;padding:11px 24px;margin:0 10px 6px 0;"
        f"background:{color};color:#ffffff;text-decoration:none;font-weight:600;border-radius:8px;"
        f"font-family:Segoe UI,Arial,sans-serif;font-size:14px'>{html.escape(label)}</a>"
    )


def desk_email_html(r):
    """Step 1 - Travel Desk: arrange the booking, then send it to the manager."""
    direct = not r.get("ManagerCode")
    then = ("This employee has <b>no reporting manager on record</b>, so confirming will approve the request."
            if direct else
            f"Once arranged, click the button to send the booking to <b>{html.escape(r['ManagerName'])}</b> "
            "for approval.")
    label = "Booking Arranged - Confirm" if direct else "Booking Arranged - Send to Manager"
    return email_html(
        ref_badge_html(r),
        "<p>Dear Sagar, Jigna and Mittal,</p>",
        f"<p>{html.escape(r['EmpName'])} ({html.escape(r['EmpCode'])}) has submitted a travel request. "
        f"Please arrange tickets and accommodation. {then}</p>",
        details_table_html(r),
        "<p style='margin:22px 0 8px'>", _email_button(label, desk_url(r["RequestId"]), "#c54b36"), "</p>",
        LINK_NOTE,
    )


def manager_email_html(r):
    """Step 2 - Manager: the Travel Desk has arranged the booking; approve or reject it."""
    rid = r["RequestId"]
    return email_html(
        ref_badge_html(r),
        f"<p>Dear {html.escape(r['ManagerName'])},</p>",
        f"<p>The Travel Desk has arranged tickets and accommodation for {html.escape(r['EmpName'])} "
        f"({html.escape(r['EmpCode'])}). Please review and approve or reject the booking. "
        "A reason is required to reject.</p>",
        booking_box_html(r),
        details_table_html(r, booking=False),
        "<p style='margin:22px 0 8px'>",
        _email_button("Approve", review_url(rid, "approve"), "#7f8e28"),
        _email_button("Reject", review_url(rid, "reject"), "#c54b36"),
        "</p>",
        LINK_NOTE,
    )


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
# Travel Desk and manager links carry their own signed token; /healthz is for monitoring; /static/ is the logo
_OPEN_PREFIXES = ("/travel-desk-confirm/", "/desk/", "/review/", "/healthz", "/static/")


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
        req = load_request(req["RequestId"]) or req  # picks up the database-assigned BookingRefID
    except Exception:  # noqa: BLE001
        log.exception("DB error during insert")
        return jsonify({"ok": False, "error": "Could not save the request."}), 500

    try:
        send_mail(
            TRAVEL_DESK_TO,
            with_ref(req, f"Booking required: {req['EmpName']} ({req['EmpCode']}) - "
                          f"{req['PlaceFrom']} to {req['PlaceTo']}"),
            desk_email_html(req),
            cc=TRAVEL_DESK_CC,
        )
    except Exception:  # noqa: BLE001
        log.exception("Failed to email the Travel Desk")
        set_status(req["RequestId"], "MailFailed")
        return jsonify({"ok": False, "error": "Could not email the Travel Desk. Please contact support."}), 502

    return jsonify({
        "ok": True,
        "requestId": str(req["RequestId"]),
        "bookingRef": booking_ref(req),
        "needsApproval": emp["NeedsApproval"],
        "message": "Request submitted. The Travel Desk has been asked to arrange your booking"
                   + (f"; {req['ManagerName']} will then approve it." if emp["NeedsApproval"] else "."),
    })


def notify_decision(r, new_status, comment):
    """Step 3 - final emails, sent separately to the employee (EmpEmail) and the Travel Desk (CC as configured).
    Approved: the employee gets the final booking (Property Name, Place/Area, Cost).
    Rejected: the employee gets the manager's reason; the Travel Desk is asked to cancel the tentative booking."""
    direct = not r.get("ManagerCode")
    approved = new_status == "Approved"
    who = f"{html.escape(r['EmpName'])} ({html.escape(r['EmpCode'])})"
    trip = f"{r['PlaceFrom']} to {r['PlaceTo']}"
    dates = f"{fmt_date(r['FromDate'])} to {fmt_date(r['ToDate'])}"
    has_booking = bool(r.get("PropertyName"))
    comment_box = reason_box_html("Manager comment", comment) if (approved and comment and not direct) else ""

    if approved:
        how = ("no manager approval was required (no reporting manager on record)" if direct
               else f"approved by {html.escape(r['ManagerName'])}")
        emp_subject = with_ref(r, f"Your trip to {r['PlaceTo']} is approved")
        emp_body = email_html(
            ref_badge_html(r),
            f"<p>Dear {html.escape(r['EmpName'])},</p>",
            f"<p>Your trip from {html.escape(r['PlaceFrom'])} to {html.escape(r['PlaceTo'])} ({dates}) "
            "is <b>approved</b>." + (" Here is your final booking:" if has_booking else "") + "</p>",
            booking_box_html(r, "Your booking") if has_booking else "",
            comment_box,
            details_table_html(r, booking=False),
        )
        desk_subject = with_ref(r, f"Approved: travel booking for {r['EmpName']} ({r['EmpCode']}) - {trip}")
        desk_body = email_html(
            ref_badge_html(r),
            "<p>Dear Sagar, Jigna and Mittal,</p>",
            f"<p>The booking for {who} is <b>approved</b>: {how}. Please confirm the tickets and accommodation.</p>",
            booking_box_html(r) if has_booking else "",
            comment_box,
            details_table_html(r, booking=False),
        )
    else:
        emp_subject = with_ref(r, "Your travel request was rejected")
        emp_body = email_html(
            ref_badge_html(r),
            f"<p>Dear {html.escape(r['EmpName'])},</p>",
            f"<p>Your travel request from {html.escape(r['PlaceFrom'])} to {html.escape(r['PlaceTo'])} ({dates}) "
            f"has been <b>rejected</b> by {html.escape(r['ManagerName'])}. "
            "The Travel Desk will cancel the tentative booking.</p>",
            reason_box_html("Reason for rejection", comment or "-"),
            details_table_html(r, booking=False),
        )
        desk_subject = with_ref(r, f"Rejected - cancel booking: {r['EmpName']} ({r['EmpCode']}) - {trip}")
        desk_body = email_html(
            ref_badge_html(r),
            "<p>Dear Sagar, Jigna and Mittal,</p>",
            f"<p>The booking for {who} has been <b>rejected</b> by {html.escape(r['ManagerName'])}. "
            "<b>Please cancel the tentative tickets and accommodation below.</b></p>",
            reason_box_html("Reason for rejection", comment or "-"),
            booking_box_html(r, "Booking to cancel") if has_booking else "",
            details_table_html(r, booking=False),
        )

    errors = []
    for to, cc, subject, body in (
        ([r["EmpEmail"]], None, emp_subject, emp_body),  # always straight to the employee
        (TRAVEL_DESK_TO, TRAVEL_DESK_CC, desk_subject, desk_body),
    ):
        try:
            send_mail(to, subject, body, cc=cc)
        except Exception as exc:  # noqa: BLE001 - still try the other recipient
            errors.append(exc)
    if errors:
        raise errors[0]


def _check_link(signer, request_id, template):
    """Validate the request id and signed token. Returns (token, error_response or None)."""
    token = request.values.get("t", "")
    try:
        uuid.UUID(request_id)
    except ValueError:
        abort(404)
    if not check_token(signer, request_id, token):
        return token, (render_template(template, state="invalid", logo_url=LOGO_URL), 403)
    return token, None


def parse_booking(form):
    """Validate the Travel Desk form. Returns (values, errors) - all three fields are required."""
    area = form.get("booking_area", "").strip()[:150]
    prop = form.get("property_name", "").strip()[:200]
    raw_price = form.get("cost", "").replace(",", "").replace("₹", "").strip()
    errors = {}
    if not area:
        errors["booking_area"] = "Enter the place/area of the stay."
    if not prop:
        errors["property_name"] = "Enter the property name."
    price = None
    try:
        price = Decimal(raw_price).quantize(Decimal("0.01"))
        if not (Decimal("0") < price < Decimal("10000000000")):
            raise InvalidOperation
    except (InvalidOperation, ValueError):
        errors["cost"] = "Enter the cost as a number, e.g. 12500."
        price = None
    return {"booking_area": area, "property_name": prop, "cost": raw_price, "price": price}, errors


@app.get("/desk/<request_id>")
def desk_legacy(request_id):
    """Links in emails sent before the booking-details form was added."""
    return redirect(url_for("travel_desk_confirm", request_id=request_id, t=request.args.get("t", "")), 301)


@app.route("/travel-desk-confirm/<request_id>", methods=["GET", "POST"])
def travel_desk_confirm(request_id):
    """Step 1 - the Travel Desk enters the booking (place/area, property, price) and sends it to the manager.
    GET only shows the form (safe for mail link scanners); POST saves and emails the manager."""
    template = "travel_desk_confirm.html"
    token, err = _check_link(_desk_signer, request_id, template)
    if err:
        return err
    r = load_request(request_id)
    if not r:
        abort(404)
    page = dict(r=r, token=token, fmt_date=fmt_date, booking_rows=booking_rows, logo_url=LOGO_URL,
                values={"booking_area": r.get("BookingArea") or "", "property_name": r.get("PropertyName") or "",
                        "cost": f"{r['Cost']:.2f}" if r.get("Cost") is not None else ""},
                errors={})

    if r["Status"] != STATUS_BOOKING:
        return render_template(template, state="already", **page)
    if request.method == "GET":
        return render_template(template, state="pending", **page)

    values, errors = parse_booking(request.form)
    page.update(values=values, errors=errors)
    if errors:
        return render_template(template, state="pending", **page), 400

    direct = not r.get("ManagerCode")  # no reporting manager (HOD / senior staff): booking confirmed = approved
    to_status = "Approved" if direct else STATUS_MANAGER
    if not save_booking(request_id, values["booking_area"], values["property_name"], values["price"],
                        to_status, comment=DIRECT_NOTE if direct else None):
        page["r"] = load_request(request_id)
        return render_template(template, state="already", **page)
    r = page["r"] = load_request(request_id)  # now includes the booking details

    if direct:
        try:
            notify_decision(r, "Approved", "")
        except Exception:  # noqa: BLE001
            log.exception("Final emails failed for %s", request_id)
        return render_template(template, state="done", **page)

    try:
        send_mail([r["ManagerEmail"]],
                  with_ref(r, f"Approval required: travel booking for {r['EmpName']} ({r['EmpCode']})"),
                  manager_email_html(r))
    except Exception:  # noqa: BLE001
        log.exception("Failed to email the manager for %s", request_id)
        advance_status(request_id, STATUS_MANAGER, STATUS_BOOKING)  # details stay saved; the desk can retry
        page["r"] = load_request(request_id)
        return render_template(template, state="pending", error=(
            "The booking was saved but could not be emailed to the manager. Please submit again."), **page), 502
    return render_template(template, state="done", **page)


@app.route("/review/<request_id>", methods=["GET", "POST"])
def review(request_id):
    """Step 2 - manager approves or rejects the booking. GET only shows the page; POST decides."""
    token, err = _check_link(_review_signer, request_id, "review.html")
    if err:
        return err
    r = load_request(request_id)  # trust the DB record only
    if not r:
        abort(404)
    page = dict(r=r, token=token, fmt_date=fmt_date, booking_rows=booking_rows, logo_url=LOGO_URL)

    if request.method == "GET":
        if r["Status"] != STATUS_MANAGER:
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
        log.exception("Final emails failed for %s", request_id)
    page["r"] = load_request(request_id)
    return render_template("review.html", state="done", **page)

if __name__ == "__main__":
    from waitress import serve
    print(f"Serving Travel Accommodation Portal on http://{HOST}:{PORT} (links use {APP_BASE_URL or 'APP_BASE_URL not set'})",
          flush=True)
    serve(app, host=HOST, port=PORT, threads=8)
