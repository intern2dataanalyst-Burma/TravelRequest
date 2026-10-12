# Travel Accommodation Portal

Internal (KMS) portal for travel and accommodation requests, hosted on **10.10.30.90** alongside SQL Server `MingalabaDB`.

**URL (office network / VPN only):** http://10.10.30.90:8080

## How it works

1. **Employee submits** the form. Name, designation and department come from `tblEmployeeMaster`; the manager from `dbo.checklist_empmanager_mapping`. Status: `Pending Booking`.
2. **Travel Desk books.** The Travel Desk (To: Jigna, Mittal, Sagar; CC: Pooja) gets an email with a **Booking Arranged - Send to Manager** button. Confirming on the portal sets `Pending Manager Approval` and emails the manager.
3. **Manager approves.** The manager's email has **Approve** and **Reject** buttons that open a review page on the portal. A reason is required to reject.
4. **Final emails** go to the employee and the Travel Desk: `Approved`, or `Rejected` with the reason (the Travel Desk is asked to cancel the tentative booking).

Employees with no reporting manager on record skip step 3: the Travel Desk's confirmation approves the request.

All emails are plain HTML sent through Microsoft Graph (no Outlook Actionable Messages). The links are signed with `SECRET_KEY` (separately for the Travel Desk and the manager) and expire after `REVIEW_LINK_MAX_AGE_DAYS` (30 by default). Opening a link changes nothing; the action is recorded only when the page is submitted. Links work only from the office network or VPN.

Database: run `deploy/migrate_001_status_length.sql` once (widens `TravelRequests.Status` to `nvarchar(30)`). Already applied on MingalabaDB on 2026-10-12.

## Deploy on 10.10.30.90

Prerequisites on the server: Python 3.10+ (installed for all users, on PATH) and *ODBC Driver 18 for SQL Server*.

1. Copy this folder to the server, e.g. `C:\Apps\TravelRequest` (git clone, or copy over RDP).
2. Put the `.env` file in that folder and replace every `<...>` placeholder.
3. In an **elevated** PowerShell, from that folder:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\deploy\install.ps1
   ```
   The script installs packages into `.venv`, checks the database connection, opens TCP 8080 in Windows Firewall (Domain/Private networks), and registers a `TravelPortal` scheduled task that starts the portal at boot and restarts it if it stops.

## Operate

| Task | Command (elevated PowerShell) |
|---|---|
| Status | `Get-ScheduledTask TravelPortal` |
| Restart (e.g. after editing `.env` or updating code) | `powershell -ExecutionPolicy Bypass -File .\deploy\restart.ps1` |
| Logs | `Get-Content .\logs\portal.log -Tail 50 -Wait` |
| Health check | `Invoke-WebRequest http://127.0.0.1:8080/healthz` |

To run it in the foreground for troubleshooting, stop the task (`Stop-ScheduledTask TravelPortal`, then end any `python.exe` still holding port 8080) and run `.\.venv\Scripts\python app.py` instead.
