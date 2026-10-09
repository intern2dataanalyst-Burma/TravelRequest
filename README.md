# Travel Accommodation Portal

Internal (KMS) portal for travel and accommodation requests, hosted on **10.10.30.90** alongside SQL Server `MingalabaDB`.

**URL (office network / VPN only):** http://10.10.30.90:8080

## How it works

1. An employee fills in the form. The app looks up their manager in `dbo.checklist_empmanager_mapping` and saves the request to `TravelRequests`.
2. The manager gets an email (sent through Microsoft Graph) with **Approve** and **Reject** buttons. They open a review page on the portal where the manager confirms. A reason is required to reject.
3. On approval, the Travel Desk (To: Jigna, Mittal, Sagar; CC: Pooja) is emailed. The employee is emailed either way.

The links in the email are signed with `SECRET_KEY` and expire after `REVIEW_LINK_MAX_AGE_DAYS` (30 by default). Opening a link changes nothing; the decision is only recorded when the manager submits the page. They work only from the office network or VPN.

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
