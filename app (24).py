"""
Travel Accommodation Portal
Flask + pyodbc (SQL Server) + MSAL/Microsoft Graph + Outlook Actionable Messages
"""
import html
import json
import logging
import os
import re
import uuid
from datetime import date, datetime

import jwt  # PyJWT
import msal
import pyodbc
import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, make_response, render_template, request
from jwt import PyJWKClient

load_dotenv()

# --------------------------------------------------------------------------
# Configuration (all from environment / App Settings)
# --------------------------------------------------------------------------
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

# Mailbox the app sends from (must also be registered as a sender for the Originator)
SENDER_MAILBOX = os.environ.get("SENDER_MAILBOX", "")

# Provider ID from the Actionable Email Developer Dashboard
ORIGINATOR_ID = os.environ.get("ORIGINATOR_ID", "REPLACE-WITH-ORIGINATOR-ID")

# Public base URL of this app, no trailing slash, e.g. https://travel-portal.azurewebsites.net
APP_BASE_URL = os.environ.get("APP_BASE_URL", "").rstrip("/")

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

# Optional login gate for the form (HTTP Basic). Leave PORTAL_PASSWORD empty to disable (local dev only).
PORTAL_USER = os.environ.get("PORTAL_USER", "")
PORTAL_PASSWORD = os.environ.get("PORTAL_PASSWORD", "")

# Brand logo shown in the header (same logo as the Medical Tracker; override with LOGO_URL, empty hides it)
LOGO_URL = os.environ.get(
    "LOGO_URL",
    "https://drive.google.com/thumbnail?id=1OpUw3MCFGLRs7GQ4xezk5ouqYTvy6yv9&sz=w200",
)

# Actionable Message token validation
AM_ISSUER = "https://substrate.office.com/sts/"
AM_OPENID_CONFIG = "https://substrate.office.com/sts/common/.well-known/openid-configuration"
AM_APP_ID = os.environ.get("AM_APP_ID", "")  # optional extra check on the token's appid claim

ROOM_CATEGORIES = {"Single Occupancy", "Double Occupancy"}
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("travel-portal")


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
    """Employee + manager from the mapping table (emails are columns in that table). None if no row."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            f"SELECT TOP 1 ManagerCode, {_q(EMP_EMAIL_COL)}, {_q(MGR_EMAIL_COL)} "
            f"FROM {_q(MAPPING_TABLE)} WHERE EmpCode = ? ORDER BY id DESC",
            emp_code,
        )
        row = cur.fetchone()
    if not row or not _clean(row[0]):
        return None
    mgr_code = _clean(row[0])
    return {
        "EmpCode": emp_code,
        "EmpName": None,              # not stored in the mapping table; the typed name is used
        "EmpEmail": _clean(row[1]),
        "ManagerCode": mgr_code,
        "ManagerName": mgr_code,      # shown as the code unless you add a ManagerName column
        "ManagerEmail": _clean(row[2]),
    }


def insert_request(req):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO TravelRequests
              (RequestId, EmpCode, EmpName, EmpEmail, ManagerCode, ManagerName, ManagerEmail,
               Designation, Department, FromDate, ToDate, PlaceFrom, PlaceTo,
               RoomCategory, PreferredLocation, TravelReason, Status)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'Pending')
            """,
            req["RequestId"], req["EmpCode"], req["EmpName"], req["EmpEmail"],
            req["ManagerCode"], req["ManagerName"], req["ManagerEmail"],
            req["Designation"], req["Department"], req["FromDate"], req["ToDate"],
            req["PlaceFrom"], req["PlaceTo"], req["RoomCategory"], req["PreferredLocation"],
            req["TravelReason"],
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
        ("Place (From)", r["PlaceFrom"]),
        ("Place (To)", r["PlaceTo"]),
        ("Reason of Travel", r.get("TravelReason") or "-"),
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
# Adaptive Card builders (v1.4)
# --------------------------------------------------------------------------
def build_approval_card(r, notice=None):
    rid = str(r["RequestId"])
    action_url = f"{APP_BASE_URL}/api/approve"

    def http_action(title, decision, style):
        return {
            "type": "Action.Http",
            "title": title,
            "style": style,
            "method": "POST",
            "url": action_url,
            "headers": [{"name": "Content-Type", "value": "application/json"}],
            "body": json.dumps({
                "requestId": rid,
                "decision": decision,
                "comment": "{{comment.value}}",
                "travelDetails": {
                    "empCode": r["EmpCode"],
                    "empName": r["EmpName"],
                    "from": fmt_date(r["FromDate"]),
                    "to": fmt_date(r["ToDate"]),
                    "placeFrom": r["PlaceFrom"],
                    "placeTo": r["PlaceTo"],
                    "roomCategory": r["RoomCategory"],
                },
            }),
        }

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "originator": ORIGINATOR_ID,
        "body": ([{"type": "TextBlock", "text": notice, "color": "Attention", "weight": "Bolder",
                   "wrap": True}] if notice else []) + [
            {"type": "TextBlock", "text": "Travel Accommodation Approval", "size": "Large",
             "weight": "Bolder", "wrap": True},
            {"type": "TextBlock", "spacing": "None", "isSubtle": True, "wrap": True,
             "text": f"{r['EmpName']} ({r['EmpCode']}) has requested travel and accommodation."},
            {"type": "FactSet", "separator": True, "facts": [
                {"title": "Emp Code", "value": r["EmpCode"]},
                {"title": "Name", "value": r["EmpName"]},
                {"title": "Designation", "value": r["Designation"]},
                {"title": "Department", "value": r["Department"]},
                {"title": "From (Date)", "value": fmt_date(r["FromDate"])},
                {"title": "To (Date)", "value": fmt_date(r["ToDate"])},
                {"title": "Place (From)", "value": r["PlaceFrom"]},
                {"title": "Place (To)", "value": r["PlaceTo"]},
                {"title": "Reason of Travel", "value": r.get("TravelReason") or "-"},
                {"title": "Category of Room", "value": r["RoomCategory"]},
                {"title": "Preferred Location", "value": r["PreferredLocation"] or "-"},
            ]},
            {"type": "Input.Text", "id": "comment", "isMultiline": True, "maxLength": 300,
             "label": "Comment (optional for Approve, REQUIRED for Reject)",
             "placeholder": "Reason for rejection / comment (shown to employee)"},
        ],
        "actions": [
            http_action("Approve", "approve", "positive"),
            http_action("Reject", "reject", "destructive"),
        ],
    }


def build_result_card(title, message, color="Good"):
    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": title, "size": "Large", "weight": "Bolder",
             "color": color, "wrap": True},
            {"type": "TextBlock", "text": message, "wrap": True},
        ],
    }


def card_email_html(r, card):
    """HTML body embedding the Adaptive Card as an Actionable Message."""
    fallback = (
        "<p style='font-family:Segoe UI,Arial,sans-serif'>"
        f"{html.escape(r['EmpName'])} has requested travel accommodation and needs your approval.</p>"
        + details_table_html(r)
    )
    card_json = json.dumps(card).replace("</", "<\\/")  # keep </script> out of the JSON
    return (
        "<html><head>"
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
        f'<script type="application/adaptivecard+json">{card_json}</script>'
        f"</head><body>{fallback}</body></html>"
    )


# --------------------------------------------------------------------------
# Actionable Message token validation
# --------------------------------------------------------------------------
_jwk_client = None


def _jwks():
    global _jwk_client
    if _jwk_client is None:
        cfg = requests.get(AM_OPENID_CONFIG, timeout=10).json()
        _jwk_client = PyJWKClient(cfg["jwks_uri"])
    return _jwk_client


def validate_action_token():
    """Validate the bearer token Outlook attaches to Action.Http calls. Returns claims or None."""
    auth = request.headers.get("Authorization", "")
    if not auth.lower().startswith("bearer "):
        return None
    token = auth[7:]
    try:
        key = _jwks().get_signing_key_from_jwt(token).key
        claims = jwt.decode(
            token, key, algorithms=["RS256"], audience=APP_BASE_URL, issuer=AM_ISSUER,
            options={"require": ["exp", "iss", "aud"]},
        )
    except Exception as exc:  # noqa: BLE001
        log.warning("Action token rejected: %s", exc)
        return None

    sender = str(claims.get("sender", "")).lower()
    if sender != SENDER_MAILBOX.lower():
        log.warning("Action token sender mismatch: %s", sender)
        return None
    if AM_APP_ID and claims.get("appid") != AM_APP_ID:
        log.warning("Action token appid mismatch")
        return None
    return claims


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
import secrets  # noqa: E402
from flask import Response  # noqa: E402

_OPEN_PATHS = ("/api/approve", "/healthz")  # webhook has its own token check; health is for Railway


@app.before_request
def require_portal_login():
    if not PORTAL_PASSWORD or request.path in _OPEN_PATHS:
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
        "managerName": emp["ManagerName"],
        "managerCode": emp["ManagerCode"],
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
        "Place (From)": place_from, "Place (To)": place_to, "Reason of Travel": travel_reason,
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
    emp["EmpName"] = emp["EmpName"] or get("emp_name")
    if not emp["EmpName"]:
        return jsonify({"ok": False, "error": "Name of Person travelling is required"}), 400
    if not emp["ManagerEmail"]:
        return jsonify({"ok": False, "error": f"No email found for manager code {emp['ManagerCode']}"}), 422
    if not emp["EmpEmail"]:
        return jsonify({"ok": False, "error": f"No email found for employee code {emp_code}"}), 422

    # Name / manager always come from the DB, never from the browser
    req = {
        "RequestId": str(uuid.uuid4()),
        **{k: emp[k] for k in ("EmpCode", "EmpName", "EmpEmail", "ManagerCode", "ManagerName", "ManagerEmail")},
        "Designation": designation, "Department": department,
        "FromDate": from_date, "ToDate": to_date,
        "PlaceFrom": place_from, "PlaceTo": place_to,
        "RoomCategory": room, "PreferredLocation": pref_loc,
        "TravelReason": travel_reason,
    }

    try:
        insert_request(req)
    except Exception:  # noqa: BLE001
        log.exception("DB error during insert")
        return jsonify({"ok": False, "error": "Could not save the request."}), 500

    try:
        card = build_approval_card(req)
        send_mail(
            [req["ManagerEmail"]],
            f"Approval required: Travel request for {req['EmpName']} ({req['EmpCode']})",
            card_email_html(req, card),
        )
    except Exception:  # noqa: BLE001
        log.exception("Failed to send approval card")
        set_status(req["RequestId"], "MailFailed")
        return jsonify({"ok": False, "error": "Could not email your manager. Please contact support."}), 502

    return jsonify({
        "ok": True,
        "requestId": req["RequestId"],
        "message": f"Request submitted. Approval email sent to {req['ManagerName']}.",
    })


@app.post("/api/approve")
def api_approve():
    claims = validate_action_token()
    if claims is None:
        return jsonify({"error": "Unauthorized"}), 401

    data = request.get_json(silent=True) or {}
    request_id = str(data.get("requestId", "")).strip()
    decision = str(data.get("decision", "")).lower()
    comment = str(data.get("comment", "") or "").strip()[:300]
    if decision not in ("approve", "reject"):
        return jsonify({"error": "Invalid decision"}), 400
    try:
        uuid.UUID(request_id)
    except ValueError:
        return jsonify({"error": "Invalid requestId"}), 400

    r = load_request(request_id)  # trust the DB record, not the payload details
    if not r:
        return jsonify({"error": "Request not found"}), 404

    # A rejection must carry a reason: re-show the same card (buttons stay active) with a notice
    if decision == "reject" and not comment and r["Status"] == "Pending":
        card = build_approval_card(
            r, notice="Reason required: please type the reason for rejection in the box, then click Reject again.")
        resp = make_response(jsonify(card), 200)
        resp.headers["CARD-UPDATE-IN-BODY"] = "true"
        return resp

    new_status = "Approved" if decision == "approve" else "Rejected"
    if not claim_decision(request_id, new_status, comment):
        current = load_request(request_id)["Status"]
        card = build_result_card("Already actioned", f"This request was already {current}.", "Warning")
    else:
        note_label = "Reason for rejection" if new_status == "Rejected" else "Manager comment"
        notes = f"<p><b>{note_label}:</b> {html.escape(comment)}</p>" if comment else ""
        try:
            if new_status == "Approved":
                send_mail(
                    TRAVEL_DESK_TO,
                    f"Travel booking required: {r['EmpName']} ({r['EmpCode']}) - {r['PlaceFrom']} to {r['PlaceTo']}",
                    "<p style='font-family:Segoe UI,Arial,sans-serif'>Dear Jigna, Mittal and Sagar,</p>"
                    f"<p>The following travel request has been <b>approved</b> by {html.escape(r['ManagerName'])}. "
                    "Please arrange ticketing and accommodation.</p>"
                    + details_table_html(r) + notes,
                    cc=TRAVEL_DESK_CC,
                )
                emp_msg = (f"<p>Dear {html.escape(r['EmpName'])},</p><p>Your travel request has been "
                           "<b>approved</b> and forwarded to the Travel Desk.</p>")
                subject = "Travel request approved"
            else:
                emp_msg = (f"<p>Dear {html.escape(r['EmpName'])},</p><p>Your travel request has been "
                           f"<b>rejected</b> by {html.escape(r['ManagerName'])}.</p>")
                subject = "Travel request rejected"
            send_mail([r["EmpEmail"]], subject,
                      "<div style='font-family:Segoe UI,Arial,sans-serif'>" + emp_msg
                      + details_table_html(r) + notes + "</div>")
        except Exception:  # noqa: BLE001
            log.exception("Post-decision email failed for %s", request_id)

        card = build_result_card(
            f"Action Completed: {new_status}",
            f"{r['EmpName']}'s travel request ({fmt_date(r['FromDate'])} to {fmt_date(r['ToDate'])}) "
            f"has been {new_status.lower()}."
            + (f" Reason: {comment}" if new_status == "Rejected" else ""),
            "Good" if new_status == "Approved" else "Attention",
        )

    resp = make_response(jsonify(card), 200)
    resp.headers["CARD-UPDATE-IN-BODY"] = "true"
    return resp


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5000)), debug=True)  # local only; Railway uses gunicorn
