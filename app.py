"""
Travel Accommodation Portal (internal network)
Flask + pyodbc (SQL Server) + MSAL/Microsoft Graph mail + signed approval links
"""
import html
import ipaddress
import json
import logging
import os
import re
import secrets
import threading
import time
import uuid
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import msal
import pyodbc
import requests
from dotenv import load_dotenv
from flask import (Flask, Response, abort, flash, jsonify, redirect, render_template, request, session,
                   url_for)
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

# The portal's own employee -> manager routing table (migration 005), maintained by HR in User Management.
# dbo.checklist_empmanager_mapping belongs to the checklist apps: it was only read once, to seed this table.
MAPPING_TABLE = os.environ.get("MAPPING_TABLE", "dbo.TravelApproverMapping")
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

# User Management (HR edits the employee -> manager mapping). Empty password = section disabled.
HR_ADMIN_PASSWORD = os.environ.get("HR_ADMIN_PASSWORD", "")
# Who may open /user-management, as comma-separated networks. Default: office/private networks only, because
# the portal itself is reachable from the internet. Set to "any" to allow everyone (HTTPS strongly advised).
HR_ALLOWED_NETWORKS = os.environ.get(
    "HR_ALLOWED_NETWORKS", "10.0.0.0/8,172.16.0.0/12,192.168.0.0/16,127.0.0.0/8,::1/128")
HR_SESSION_MINUTES = int(os.environ.get("HR_SESSION_MINUTES", "30"))

ROOM_CATEGORIES = {"Single Occupancy", "Double Occupancy"}

# Workflow: employee submits -> Travel Desk books -> manager approves/rejects
STATUS_BOOKING = "Pending Booking"            # waiting for the Travel Desk to arrange the booking
STATUS_MANAGER = "Pending Manager Approval"   # booking arranged, waiting for the manager
DIRECT_NOTE = "No reporting manager on record: approved when the Travel Desk confirmed the booking"
MAX_BOOKING_OPTIONS = 5  # options the Travel Desk can offer the manager
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

# Signed session cookie for the User Management unlock (HttpOnly; not readable by page scripts)
app.secret_key = SECRET_KEY
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  SESSION_COOKIE_NAME="tp_session")


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


def save_booking(request_id, options, to_status, comment=None):
    """Store the Travel Desk's booking options (JSON) and move Pending Booking -> to_status atomically.
    to_status is Pending Manager Approval, or Approved when there is no reporting manager - then the single
    option is also written to BookingArea/PropertyName/Cost as the final booking. Returns True if this call won."""
    approved = to_status == "Approved"
    sets, args = ["BookingOptions = ?", "Status = ?"], [options_json(options), to_status]
    if approved:
        final = options[0]
        sets += ["BookingArea = ?", "PropertyName = ?", "Cost = ?", "SelectedOption = 1",
                 "ActionedAt = SYSUTCDATETIME()", "ActionComment = ?"]
        args += [final["area"], final["property"], final["cost"], comment]
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"UPDATE TravelRequests SET {', '.join(sets)} WHERE RequestId = ? AND Status = ?",
                    *args, request_id, STATUS_BOOKING)
        won = cur.rowcount == 1
        conn.commit()
        return won


def claim_decision(request_id, status, comment, from_status=None, option=None, option_no=None):
    """Atomically move a request awaiting a decision to Approved/Rejected. On approval, the manager's chosen
    booking option (1-based option_no) becomes the final BookingArea/PropertyName/Cost in the same update.
    Returns True if this call won."""
    chosen = ", BookingArea = ?, PropertyName = ?, Cost = ?, SelectedOption = ?" if option else ""
    chosen_args = [option["area"], option["property"], option["cost"], option_no] if option else []
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE TravelRequests SET Status = ?, ActionedAt = SYSUTCDATETIME(), ActionComment = ?, "
            f"RejectionReason = ?{chosen} WHERE RequestId = ? AND Status = ?",
            status, comment, (comment if status == "Rejected" else None), *chosen_args, request_id,
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


def fmt_duration(from_date, to_date):
    """'3 nights, 4 days' - same wording as the form's Trip summary."""
    nights = (to_date - from_date).days
    days = nights + 1
    return f"{nights} night{'' if nights == 1 else 's'}, {days} day{'' if days == 1 else 's'}"


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
        ("Duration", fmt_duration(r["FromDate"], r["ToDate"])),
        ("Travelling From", r["PlaceFrom"]),
        ("Travelling To", r["PlaceTo"]),
        ("Reason of Travel", r.get("ReasonOfTravel") or "-"),
        ("Category of Room", r["RoomCategory"]),
        ("Preferred Location", r["PreferredLocation"] or "-"),
    ]
    if booking and r.get("PropertyName"):
        rows += booking_rows(r)
    last = len(rows) - 1
    body = "".join(
        "<tr>"
        f"<td width='40%' style='width:40%;padding:10px 14px;background:#f6efe6;color:#6b5a4e;font-size:13px;"
        f"font-weight:600;vertical-align:top;{'' if i == last else 'border-bottom:1px solid #ece2d6;'}'>"
        f"{html.escape(k)}</td>"
        f"<td style='padding:10px 14px;color:#3b2f2a;font-size:14px;font-weight:600;vertical-align:top;"
        f"word-break:break-word;{'' if i == last else 'border-bottom:1px solid #ece2d6;'}'>{html.escape(str(v))}</td>"
        "</tr>"
        for i, (k, v) in enumerate(rows)
    )
    # Rounded warm card; overflow:hidden keeps the tinted label column inside the rounded corners
    return (
        "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' "
        "style='width:100%;max-width:600px;border-collapse:separate;border-spacing:0;background:#fdfaf6;"
        "border:1px solid #e6ddd3;border-radius:12px;overflow:hidden;font-family:Segoe UI,Arial,sans-serif'>"
        f"{body}</table>"
    )


def booking_box_html(r, title="Booking arranged by the Travel Desk", rows=None, emphasis=False):
    """A booking (Place/Area, Property Name, Cost) in a soft, brand-tinted box.
    rows defaults to the final booking; emphasis=True gives the approved option a stronger border."""
    rows = "".join(
        "<tr>"
        f"<td width='40%' style='width:40%;padding:6px 12px 6px 0;color:#8a5a4c;font-size:13px'>{html.escape(k)}</td>"
        "<td style='padding:6px 0;color:#1d1d1d;font-size:15px;font-weight:700;word-break:break-word'>"
        f"{html.escape(v)}</td></tr>"
        for k, v in (rows if rows is not None else booking_rows(r))
    )
    border = "border:2px solid #c54b36" if emphasis else "border:1px solid #f3dccf"
    return (
        f"<table {_BOX.format('background:#fdf3ec;' + border)}>"
        "<tr><td style='padding:14px 18px;font-family:Segoe UI,Arial,sans-serif'>"
        "<div style='font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#c54b36;"
        f"margin-bottom:6px'>{html.escape(title)}</div>"
        f"<table {_TABLE}>{rows}</table>"
        "</td></tr></table>"
    )


def options_html(r, heading_note=""):
    """Every option the Travel Desk offered, one box each, labelled Option 1, Option 2, ..."""
    opts = booking_options(r)
    return "".join(booking_box_html(r, f"Option {i}{heading_note}", option_rows(o)) for i, o in enumerate(opts, 1))


def approved_box_html(r, title):
    """The option the manager approved, highlighted."""
    n, total = selected_option_no(r), len(booking_options(r))
    label = f"{title} (Option {n})" if n and total > 1 else title
    return booking_box_html(r, label, emphasis=True)


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
    """The final (approved) booking as (label, value) pairs."""
    return [
        ("Place/Area", r.get("BookingArea") or "-"),
        ("Property Name", r.get("PropertyName") or "-"),
        ("Cost", fmt_inr(r["Cost"]) if r.get("Cost") is not None else "-"),
    ]


def options_json(options):
    return json.dumps([{"area": o["area"], "property": o["property"], "cost": f"{o['cost']:.2f}"} for o in options],
                      ensure_ascii=False)


def booking_options(r):
    """All options the Travel Desk offered, as [{'area', 'property', 'cost': Decimal}].
    Requests saved before multiple options existed have just their single booking."""
    raw = r.get("BookingOptions")
    if raw:
        try:
            return [{"area": o["area"], "property": o["property"], "cost": Decimal(o["cost"])} for o in json.loads(raw)]
        except (ValueError, KeyError, TypeError, InvalidOperation):
            log.warning("Unreadable BookingOptions for %s", r.get("RequestId"))
    if r.get("PropertyName"):
        return [{"area": r.get("BookingArea") or "", "property": r["PropertyName"], "cost": r.get("Cost") or Decimal(0)}]
    return []


def option_rows(o):
    return [("Place/Area", o["area"] or "-"), ("Property Name", o["property"] or "-"), ("Cost", fmt_inr(o["cost"]))]


def selected_option_no(r):
    """1-based number of the option the manager approved (1 for single-option and older requests)."""
    return r.get("SelectedOption") or (1 if r.get("PropertyName") else None)


app.jinja_env.globals.update(booking_ref=booking_ref, fmt_duration=fmt_duration, booking_options=booking_options,
                             option_rows=option_rows, selected_option_no=selected_option_no,
                             max_booking_options=MAX_BOOKING_OPTIONS)  # same values as the emails


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
    then = ("Once arranged, click <b>Enter Booking Details</b> to add the booking. This employee has "
            "<b>no reporting manager on record</b>, so saving it approves the request."
            if direct else
            "Once arranged, click <b>Enter Booking Details</b> to add the booking and send it to "
            f"<b>{html.escape(r['ManagerName'])}</b> for approval.")
    label = "Enter Booking Details"
    return email_html(
        ref_badge_html(r),
        "<p>Dear Travel Desk Team,</p>",
        f"<p>{html.escape(r['EmpName'])} ({html.escape(r['EmpCode'])}) has submitted a travel request. "
        f"Please arrange tickets and accommodation. {then}</p>",
        details_table_html(r),
        "<p style='margin:22px 0 8px'>", _email_button(label, desk_url(r["RequestId"]), "#c54b36"), "</p>",
        LINK_NOTE,
    )


def manager_email_html(r):
    """Step 2 - Manager: the Travel Desk has arranged the booking; approve or reject it."""
    rid = r["RequestId"]
    n = len(booking_options(r))
    ask = (f"The Travel Desk has found <b>{n} accommodation options</b>. Please choose one to approve, "
           "or reject the request." if n > 1 else "Please review and approve or reject the booking.")
    return email_html(
        ref_badge_html(r),
        f"<p>Dear {html.escape(r['ManagerName'])},</p>",
        f"<p>The Travel Desk has arranged tickets and accommodation for {html.escape(r['EmpName'])} "
        f"({html.escape(r['EmpCode'])}). {ask} A reason is required to reject.</p>",
        options_html(r),
        details_table_html(r, booking=False),
        "<p style='margin:22px 0 8px'>",
        _email_button("Choose & Approve" if n > 1 else "Approve", review_url(rid, "approve"), "#7f8e28"),
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
    return render_template("index.html", today=today, logo_url=LOGO_URL, show_hr_link=True)


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
    multi = len(booking_options(r)) > 1
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
            approved_box_html(r, "Your booking") if has_booking else "",
            comment_box,
            details_table_html(r, booking=False),
        )
        desk_subject = with_ref(r, f"Approved: travel booking for {r['EmpName']} ({r['EmpCode']}) - {trip}")
        desk_body = email_html(
            ref_badge_html(r),
            "<p>Dear Travel Desk Team,</p>",
            f"<p>The booking for {who} is <b>approved</b>: {how}. "
            + (f"The manager chose <b>Option {selected_option_no(r)}</b>; please confirm it and release the "
               "other options." if multi else "Please confirm the tickets and accommodation.") + "</p>",
            approved_box_html(r, "Approved booking") if has_booking else "",
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
            "<p>Dear Travel Desk Team,</p>",
            f"<p>The booking for {who} has been <b>rejected</b> by {html.escape(r['ManagerName'])}. "
            "<b>Please cancel the tentative tickets and accommodation below.</b></p>",
            reason_box_html("Reason for rejection", comment or "-"),
            options_html(r, " - cancel"),
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


def _parse_cost(raw):
    try:
        cost = Decimal(raw.replace(",", "").replace("\u20b9", "").strip()).quantize(Decimal("0.01"))
        return cost if Decimal("0") < cost < Decimal("10000000000") else None
    except (InvalidOperation, ValueError):
        return None


def parse_booking_options(form, allow_multiple=True):
    """Validate the Travel Desk form. Option 1 is required; every extra option row must be complete.
    Returns (rows for re-display, options, errors) where errors maps 'field-index' -> message."""
    areas, props, costs = form.getlist("booking_area"), form.getlist("property_name"), form.getlist("cost")
    n = max(len(areas), len(props), len(costs), 1)
    rows, options, errors = [], [], {}
    for i in range(n):
        area = (areas[i] if i < len(areas) else "").strip()[:150]
        prop = (props[i] if i < len(props) else "").strip()[:200]
        raw = (costs[i] if i < len(costs) else "").strip()[:20]
        rows.append({"booking_area": area, "property_name": prop, "cost": raw})
        cost = _parse_cost(raw)
        if not area:
            errors[f"booking_area-{i}"] = "Enter the place/area of the stay."
        if not prop:
            errors[f"property_name-{i}"] = "Enter the property name."
        if cost is None:
            errors[f"cost-{i}"] = "Enter the cost as a number, e.g. 12500."
        options.append({"area": area, "property": prop, "cost": cost})
    if n > (MAX_BOOKING_OPTIONS if allow_multiple else 1):
        errors["form"] = (f"You can offer at most {MAX_BOOKING_OPTIONS} options." if allow_multiple else
                          "This employee has no reporting manager, so only one booking can be entered.")
    return rows, options, errors


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
    saved = booking_options(r)  # e.g. after a failed manager email, show what was entered
    page = dict(r=r, token=token, fmt_date=fmt_date, booking_rows=booking_rows, logo_url=LOGO_URL, errors={},
                rows=[{"booking_area": o["area"], "property_name": o["property"], "cost": f"{o['cost']:.2f}"}
                      for o in saved] or [{"booking_area": "", "property_name": "", "cost": ""}])

    if r["Status"] != STATUS_BOOKING:
        return render_template(template, state="already", **page)
    if request.method == "GET":
        return render_template(template, state="pending", **page)

    direct = not r.get("ManagerCode")  # no reporting manager (HOD / senior staff): booking confirmed = approved
    rows, options, errors = parse_booking_options(request.form, allow_multiple=not direct)
    page.update(rows=rows, errors=errors)
    if errors:
        return render_template(template, state="pending", **page), 400

    to_status = "Approved" if direct else STATUS_MANAGER
    if not save_booking(request_id, options, to_status, comment=DIRECT_NOTE if direct else None):
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
        return render_template("review.html", state="pending", decision=decision, comment="", picked="", **page)

    decision = request.form.get("decision", "").lower()
    comment = request.form.get("comment", "").strip()[:300]
    picked = request.form.get("option", "")
    page.update(picked=picked)
    if decision not in ("approve", "reject"):
        return render_template("review.html", state="pending", decision="", comment=comment,
                               error="Choose Approve or Reject.", **page), 400
    if decision == "reject" and not comment:
        return render_template("review.html", state="pending", decision="reject", comment="",
                               error="A reason is required to reject this request.", **page), 400

    option = option_no = None
    if decision == "approve":
        options = booking_options(r)
        if len(options) == 1 and not picked:
            picked = "1"  # a single option is pre-selected
        if options:
            if not (picked.isdigit() and 1 <= int(picked) <= len(options)):
                return render_template("review.html", state="pending", decision="approve", comment=comment,
                                       error="Choose which booking option to approve.", **page), 400
            option_no = int(picked)
            option = options[option_no - 1]

    new_status = "Approved" if decision == "approve" else "Rejected"
    if not claim_decision(request_id, new_status, comment, option=option, option_no=option_no):
        page["r"] = load_request(request_id)
        return render_template("review.html", state="already", **page)

    r = page["r"] = load_request(request_id)  # includes the chosen option for the final emails
    try:
        notify_decision(r, new_status, comment)
    except Exception:  # noqa: BLE001
        log.exception("Final emails failed for %s", request_id)
    return render_template("review.html", state="done", **page)


# --------------------------------------------------------------------------
# User Management - HR maintains the employee -> manager mapping that routes approval emails
# --------------------------------------------------------------------------
_CODE_RE = re.compile(r"^[A-Za-z0-9\-_/.]{1,50}$")
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
_hr_failures = {}            # ip -> [timestamps of failed password attempts]
_hr_failures_lock = threading.Lock()
HR_MAX_FAILURES, HR_LOCKOUT_SECONDS = 5, 15 * 60


def _hr_networks():
    if HR_ALLOWED_NETWORKS.strip().lower() == "any":
        return None
    nets = []
    for part in HR_ALLOWED_NETWORKS.split(","):
        try:
            nets.append(ipaddress.ip_network(part.strip(), strict=False))
        except ValueError:
            log.warning("Ignoring invalid HR_ALLOWED_NETWORKS entry: %r", part)
    return nets


def hr_network_allowed(ip):
    nets = _hr_networks()
    if nets is None:
        return True
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return False
    return any(addr in n for n in nets)


def _hr_locked_out(ip):
    now = time.time()
    with _hr_failures_lock:
        recent = [t for t in _hr_failures.get(ip, []) if now - t < HR_LOCKOUT_SECONDS]
        _hr_failures[ip] = recent
        return len(recent) >= HR_MAX_FAILURES


def _hr_record_failure(ip):
    with _hr_failures_lock:
        _hr_failures.setdefault(ip, []).append(time.time())


def hr_unlocked():
    """True while this browser's HR session is valid (sliding HR_SESSION_MINUTES idle timeout)."""
    last = session.get("hr_unlocked_at")
    if not last or time.time() - last > HR_SESSION_MINUTES * 60:
        session.pop("hr_unlocked_at", None)
        return False
    session["hr_unlocked_at"] = time.time()
    return True


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


def csrf_ok():
    sent = request.form.get("csrf", "")
    return bool(sent) and secrets.compare_digest(sent, session.get("csrf", ""))


app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def guard_user_management():
    if not request.path.startswith("/user-management"):
        return None
    if not hr_network_allowed(request.remote_addr or ""):
        log.warning("User Management blocked for %s (not in HR_ALLOWED_NETWORKS)", request.remote_addr)
        return render_template("hr_login.html", state="network", logo_url=LOGO_URL, show_hr_link=True), 403
    if not HR_ADMIN_PASSWORD:
        return render_template("hr_login.html", state="disabled", logo_url=LOGO_URL, show_hr_link=True), 503
    if request.endpoint in ("hr_login", "static") or hr_unlocked():
        return None
    if request.path.startswith("/user-management/api/"):
        return jsonify({"error": "locked"}), 401
    return redirect(url_for("hr_login"))


# ---- database
def list_mappings():
    """Every mapping row with names from the HR master (codes not in the master show no name)."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT m.id, m.EmpCode, e.EmployeeName, e.Status, {_q(EMP_EMAIL_COL)}, m.ManagerCode, mg.EmployeeName,
                   {_q(MGR_EMAIL_COL)}
            FROM {_q(MAPPING_TABLE)} m
            LEFT JOIN {_q(EMPLOYEE_TABLE)} e ON e.EmpCode = m.EmpCode
            LEFT JOIN {_q(EMPLOYEE_TABLE)} mg ON mg.EmpCode = m.ManagerCode
            ORDER BY m.EmpCode
            """
        )
        rows = cur.fetchall()
    return [{"id": int(r[0]), "emp_code": _clean(r[1]) or "", "emp_name": _clean(r[2]), "emp_status": _clean(r[3]),
             "emp_email": _clean(r[4]) or "", "mgr_code": _clean(r[5]) or "", "mgr_name": _clean(r[6]),
             "mgr_email": _clean(r[7]) or ""} for r in rows]


def get_mapping(mapping_id):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT id, EmpCode, ManagerCode, {_q(EMP_EMAIL_COL)}, {_q(MGR_EMAIL_COL)} "
                    f"FROM {_q(MAPPING_TABLE)} WHERE id = ?", mapping_id)
        r = cur.fetchone()
    if not r:
        return None
    return {"id": int(r[0]), "emp_code": _clean(r[1]) or "", "mgr_code": _clean(r[2]) or "",
            "emp_email": _clean(r[3]) or "", "mgr_email": _clean(r[4]) or ""}


def mapping_id_for(emp_code):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT TOP 1 id FROM {_q(MAPPING_TABLE)} WHERE EmpCode = ? ORDER BY id DESC", emp_code)
        r = cur.fetchone()
    return int(r[0]) if r else None


class SharedTableError(RuntimeError):
    """Raised instead of writing to a table that belongs to another application."""


def _writable_mapping_table():
    """The checklist apps own dbo.checklist_* tables; the portal must never write to them."""
    if "checklist_" in MAPPING_TABLE.lower():
        raise SharedTableError(f"{MAPPING_TABLE} is shared with the checklist apps; set MAPPING_TABLE to "
                               "dbo.TravelApproverMapping (migration 005)")
    return _q(MAPPING_TABLE)


def save_mapping(m, mapping_id=None):
    """Insert (mapping_id None) or update one mapping row in the portal's own table. Blank values become NULL."""
    table = _writable_mapping_table()
    vals = [m["emp_code"], m["mgr_code"] or None, m["emp_email"] or None, m["mgr_email"] or None]
    with get_conn() as conn:
        cur = conn.cursor()
        if mapping_id is None:
            cur.execute(f"INSERT INTO {table} (EmpCode, ManagerCode, {_q(EMP_EMAIL_COL)}, "
                        f"{_q(MGR_EMAIL_COL)}) VALUES (?, ?, ?, ?)", *vals)
            ok = True
        else:
            cur.execute(f"UPDATE {table} SET EmpCode = ?, ManagerCode = ?, {_q(EMP_EMAIL_COL)} = ?, "
                        f"{_q(MGR_EMAIL_COL)} = ?, UpdatedAt = SYSUTCDATETIME() WHERE id = ?", *vals, mapping_id)
            ok = cur.rowcount == 1
        conn.commit()
    return ok


def delete_mapping(mapping_id):
    table = _writable_mapping_table()
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"DELETE FROM {table} WHERE id = ?", mapping_id)
        ok = cur.rowcount == 1
        conn.commit()
    return ok


def master_lookup(code):
    """Name / email / status for a code from the HR master, or None."""
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT TOP 1 EmployeeName, EmployeeEmail, Status FROM {_q(EMPLOYEE_TABLE)} WHERE EmpCode = ?",
                    code)
        r = cur.fetchone()
    return {"name": _clean(r[0]), "email": _clean(r[1]), "status": _clean(r[2])} if r else None


def validate_mapping(form, mapping_id=None):
    """Returns (values, errors)."""
    v = {k: form.get(k, "").strip() for k in ("emp_code", "emp_email", "mgr_code", "mgr_email")}
    errors = {}
    if not _CODE_RE.match(v["emp_code"]):
        errors["emp_code"] = "Enter the employee code (letters, digits, - _ / . only)."
    if v["mgr_code"] and not _CODE_RE.match(v["mgr_code"]):
        errors["mgr_code"] = "Manager code can use letters, digits, - _ / . only."
    if v["mgr_code"] and v["mgr_code"].lower() == v["emp_code"].lower():
        errors["mgr_code"] = "An employee cannot be their own manager. Leave it blank for a HOD."
    for key in ("emp_email", "mgr_email"):
        if v[key] and (len(v[key]) > 256 or not _EMAIL_RE.match(v[key])):
            errors[key] = "Enter a valid email address, or leave it blank."
    if v["mgr_code"] and not v["mgr_email"]:
        errors["mgr_email"] = "The manager's email is needed to send approval requests."
    if "emp_code" not in errors:
        existing = mapping_id_for(v["emp_code"])
        if existing is not None and existing != mapping_id:
            errors["emp_code"] = f"{v['emp_code']} already has a mapping. Edit that row instead."
    return v, errors


def _audit(action, before=None, after=None):
    log.info("USER-MGMT %s from %s | before=%s | after=%s", action, request.remote_addr, before, after)


# ---- routes
@app.route("/user-management/login", methods=["GET", "POST"])
def hr_login():
    if hr_unlocked():
        return redirect(url_for("hr_dashboard"))
    ip = request.remote_addr or ""
    page = dict(logo_url=LOGO_URL, show_hr_link=True, hr_active=True)
    if _hr_locked_out(ip):
        return render_template("hr_login.html", state="locked", **page), 429
    if request.method == "GET":
        return render_template("hr_login.html", state="login", **page)
    if not csrf_ok():
        return render_template("hr_login.html", state="login", error="Your session expired. Please try again.",
                               **page), 400
    if secrets.compare_digest(request.form.get("password", "").encode(), HR_ADMIN_PASSWORD.encode()):
        session.clear()  # fresh session on unlock
        session["hr_unlocked_at"] = time.time()
        csrf_token()
        log.info("USER-MGMT unlocked from %s", ip)
        return redirect(url_for("hr_dashboard"))
    _hr_record_failure(ip)
    log.warning("USER-MGMT wrong password from %s", ip)
    if _hr_locked_out(ip):
        return render_template("hr_login.html", state="locked", **page), 429
    return render_template("hr_login.html", state="login", error="Incorrect password.", **page), 401


@app.post("/user-management/logout")
def hr_logout():
    if csrf_ok():
        session.pop("hr_unlocked_at", None)
    return redirect(url_for("hr_login"))


@app.get("/user-management")
def hr_dashboard(form=None, form_errors=None, status=200):
    try:
        rows = list_mappings()
    except Exception:  # noqa: BLE001
        log.exception("Could not load mappings")
        rows, form_errors = [], {"_": "Could not load the mappings from the database."}
    stats = {
        "total": len(rows),
        "no_manager": sum(1 for r in rows if not r["mgr_code"]),
        "missing_mgr_email": sum(1 for r in rows if r["mgr_code"] and not r["mgr_email"]),
        "missing_emp_email": sum(1 for r in rows if not r["emp_email"]),
    }
    return render_template("hr_users.html", rows=rows, stats=stats, form=form or {}, form_errors=form_errors or {},
                           logo_url=LOGO_URL, show_hr_link=True, hr_active=True,
                           session_minutes=HR_SESSION_MINUTES), status


@app.post("/user-management/save")
def hr_save():
    if not csrf_ok():
        flash("Your session expired. Please try again.", "bad")
        return redirect(url_for("hr_dashboard"))
    raw_id = request.form.get("id", "").strip()
    mapping_id = int(raw_id) if raw_id.isdigit() else None
    before = get_mapping(mapping_id) if mapping_id else None
    if mapping_id and not before:
        flash("That mapping no longer exists.", "bad")
        return redirect(url_for("hr_dashboard"))
    values, errors = validate_mapping(request.form, mapping_id)
    if errors:
        return hr_dashboard(form={**values, "id": raw_id}, form_errors=errors, status=400)
    try:
        save_mapping(values, mapping_id)
    except SharedTableError:
        log.exception("Refused to write the mapping")
        flash("Saving is disabled: the portal is configured to use a table shared with other apps. "
              "Ask IT to set MAPPING_TABLE=dbo.TravelApproverMapping.", "bad")
        return redirect(url_for("hr_dashboard"))
    except pyodbc.IntegrityError:  # another HR user added the same Emp Code a moment ago
        return hr_dashboard(form={**values, "id": raw_id},
                            form_errors={"emp_code": f"{values['emp_code']} already has a mapping. Edit that row instead."},
                            status=400)
    _audit("UPDATE" if mapping_id else "ADD", before, values)
    who = values["emp_code"]
    flash(f"Mapping for {who} {'updated' if mapping_id else 'added'}. New travel requests use it immediately.", "ok")
    return redirect(url_for("hr_dashboard", q=who))


@app.post("/user-management/delete")
def hr_delete():
    if not csrf_ok():
        flash("Your session expired. Please try again.", "bad")
        return redirect(url_for("hr_dashboard"))
    raw_id = request.form.get("id", "")
    before = get_mapping(int(raw_id)) if raw_id.isdigit() else None
    try:
        deleted = bool(before) and delete_mapping(before["id"])
    except SharedTableError:
        log.exception("Refused to delete from the mapping table")
        flash("Deleting is disabled: the portal is configured to use a table shared with other apps. "
              "Ask IT to set MAPPING_TABLE=dbo.TravelApproverMapping.", "bad")
        return redirect(url_for("hr_dashboard"))
    if not deleted:
        flash("That mapping no longer exists.", "bad")
        return redirect(url_for("hr_dashboard"))
    _audit("DELETE", before)
    flash(f"Mapping for {before['emp_code']} deleted.", "ok")
    return redirect(url_for("hr_dashboard"))


@app.get("/user-management/api/lookup/<code>")
def hr_lookup(code):
    """Name/email from the HR master while HR types a code in the form."""
    if not _CODE_RE.match(code):
        return jsonify({"found": False}), 400
    m = master_lookup(code)
    return jsonify({"found": bool(m), **(m or {})})


if __name__ == "__main__":
    from waitress import serve
    print(f"Serving Travel Accommodation Portal on http://{HOST}:{PORT} (links use {APP_BASE_URL or 'APP_BASE_URL not set'})",
          flush=True)
    serve(app, host=HOST, port=PORT, threads=8)
