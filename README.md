# Travel Accommodation Portal

Internal (KMS) portal for travel and accommodation requests, hosted on **10.10.30.90** alongside SQL Server `MingalabaDB`.

**URL:** `http://<Public-IP>:9000` from anywhere (the company router forwards port 9000 to 10.10.30.90), or `http://10.10.30.90:9000` in the office. Email links use `APP_BASE_URL` from `.env`.

## How it works

1. **Employee submits** the form. Name, designation and department come from `tblEmployeeMaster`; the manager from `dbo.checklist_empmanager_mapping`. Status: `Pending Booking`.
2. **Travel Desk books.** The Travel Desk (To: Jigna, Mittal, Sagar; CC: Pooja) gets an email with an **Enter Booking Details** button. It opens a form where they enter one or more accommodation options (Place/Area, Property Name, Cost; up to 5); **Save & Send for Approval** sets `Pending Manager Approval` and emails the manager every option.
3. **Manager approves.** The manager's email lists the options and has **Approve** and **Reject** buttons that open a review page on the portal. To approve, the manager picks one option (pre-selected when there is only one); the chosen option becomes the final booking. A reason is required to reject.
4. **Final emails** go to the employee and the Travel Desk: `Approved`, or `Rejected` with the reason (the Travel Desk is asked to cancel the tentative booking).

Employees with no reporting manager on record skip step 3: the Travel Desk's confirmation approves the request.

All emails are plain HTML sent through Microsoft Graph (no Outlook Actionable Messages). The links are signed with `SECRET_KEY` (separately for the Travel Desk and the manager) and expire after `REVIEW_LINK_MAX_AGE_DAYS` (30 by default). Opening a link changes nothing; the action is recorded only when the page is submitted. Links work only from the office network or VPN.

Database migrations (run once each, in order; all applied on MingalabaDB, 001-003 on 2026-10-12 and 004 on 2026-10-13):
- `deploy/migrate_001_status_length.sql` widens `TravelRequests.Status` to `nvarchar(30)`.
- `deploy/migrate_002_booking_details.sql` adds `BookingArea`, `PropertyName` and `Cost` (the Travel Desk's booking details).
- `deploy/migrate_003_booking_ref.sql` adds `BookingRefID` (auto-numbered from 10001), shown as "Booking Ref: #10001" in emails.
- `deploy/migrate_004_booking_options.sql` adds `BookingOptions` (JSON list of the Travel Desk's options) and `SelectedOption` (the one the manager approved).

## User Management (HR)

The **User Management** tab lets HR add, edit and delete rows in `dbo.checklist_empmanager_mapping`, the table that decides who approves each employee's travel. Changes apply to new requests immediately.

- Unlock with `HR_ADMIN_PASSWORD` from `.env`. The section locks again after `HR_SESSION_MINUTES` (30) of inactivity, or with **Lock**.
- 5 wrong passwords from one address pause sign-in for 15 minutes.
- By default it opens only from office/private networks (`HR_ALLOWED_NETWORKS`), even though the portal itself is public. Set it to `any` only once the portal runs on HTTPS.
- Leaving the manager code blank marks the employee as a HOD (no manager approval). Deleting a row has the same effect.
- Every add/edit/delete is written to `logs\portal.log` with the before/after values and the sender's IP.
- The same table may be used by other checklist apps on MingalabaDB; edits affect them too.

## Security note (public IP over HTTP)

The portal is reachable from the internet on plain `http://`. Passwords (portal login, HR password) and approval links travel unencrypted, so anyone on the network path (public Wi-Fi, mobile carrier) can read them. Put HTTPS in front before wide use: e.g. a company domain with a certificate on IIS (reverse proxy to port 9000), or ask IT for their preferred TLS setup. Then set `APP_BASE_URL` to the `https://` address.

## Deploy on 10.10.30.90

Prerequisites on the server: Python 3.10+ (installed for all users, on PATH) and *ODBC Driver 18 for SQL Server*.

1. Copy this folder to the server, e.g. `C:\Apps\TravelRequest` (git clone, or copy over RDP).
2. Put the `.env` file in that folder and replace every `<...>` placeholder.
3. In an **elevated** PowerShell, from that folder:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\deploy\install.ps1
   ```
   The script installs packages into `.venv`, checks the database connection, opens the `PORT` from `.env` (9000) in Windows Firewall (Domain/Private networks), and registers a `TravelPortal` scheduled task that starts the portal at boot and restarts it if it stops.

## Operate

| Task | Command (elevated PowerShell) |
|---|---|
| Status | `Get-ScheduledTask TravelPortal` |
| Restart (e.g. after editing `.env` or updating code) | `powershell -ExecutionPolicy Bypass -File .\deploy\restart.ps1` |
| Logs | `Get-Content .\logs\portal.log -Tail 50 -Wait` |
| Health check | `Invoke-WebRequest http://127.0.0.1:9000/healthz` |

To run it in the foreground for troubleshooting, stop the task (`Stop-ScheduledTask TravelPortal`, then end any `python.exe` still holding port 9000) and run `.\.venv\Scripts\python app.py` instead.
