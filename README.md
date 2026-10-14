# Travel Accommodation Portal

Internal (KMS) portal for travel and accommodation requests, hosted on **10.10.30.90** alongside SQL Server `MingalabaDB`.

**URL:** http://mingalaba.hungerpangs.com:9000/dashboard/travel/ (the `/dashboard/` root is left free for a future central menu). Email links are built from `APP_BASE_URL` in `.env`.

## How it works

1. **Employee submits** a hotel booking request. Name, designation and department come from `tblEmployeeMaster`; the manager from the portal's own routing table `dbo.TravelApproverMapping` (maintained by HR in User Management). Section 3 "Occupancy" has the room category, preferred location and an optional preferred hotel. Status: `Pending Booking`.
2. **Admin Team proposes.** The Admin Team (To: Jigna, Mittal, Sagar; CC: Pooja) gets an email with an **Enter Booking Details** button. It opens a form with one or more choices (up to 5): Place/Area, Hotel, Total Cost of stay with GST (approx), plus **Admin Team only** fields: Type of occupancy, Refund status (Refundable / Non-refundable / Partial refundable) and optional Remarks. **Save & Send for Approval** sets `Pending Manager Approval` and emails the manager.
3. **Manager approves.** The manager never sees the Admin Team only fields.
   - One choice: the email shows the booking with no "Choice 1" label. **Approve** opens a minimal confirm page (hotel, cost, the non-refundable tick-box and one **Confirm Approval** button). Opening the link changes nothing, so enterprise link scanners cannot approve by accident.
   - Several choices: **Choose & Approve** opens the review page, where the manager picks Choice 1, 2, ... and must tick "I acknowledge that once approved, this booking is non-refundable." There is no comment box for approvals.
   - **Reject** always requires a reason.
   Approval sets `Pending Final Booking`.
4. **Admin Team books.** They get an "Approved - final booking needed" email with **Upload Booking Voucher**. After booking, they upload the voucher (PDF, max 3 MB); the portal emails it as an attachment to the employee with the manager in CC and sets `Booked`.
5. On rejection, the employee gets the reason and the Admin Team is asked to cancel the tentative booking.

Employees with no reporting manager on record skip step 3: the Admin Team's confirmation approves the request and goes straight to step 4.

All emails are plain HTML sent through Microsoft Graph (no Outlook Actionable Messages). The links are signed with `SECRET_KEY` (separate keys for the Admin Team booking form, the manager review and the voucher upload) and expire after `REVIEW_LINK_MAX_AGE_DAYS` (30 by default). Opening a link never changes anything; the action is recorded only when the page's button is pressed.

(The `.env` settings `TRAVEL_DESK_TO` / `TRAVEL_DESK_CC` keep their old names and hold the Admin Team addresses.)

Database migrations (run once each, in order; 001-003 applied on 2026-10-12, 004-005 on 2026-10-13):
- `deploy/migrate_001_status_length.sql` widens `TravelRequests.Status` to `nvarchar(30)`.
- `deploy/migrate_002_booking_details.sql` adds `BookingArea`, `PropertyName` and `Cost` (the approved booking: Place/Area, Hotel, Total Cost).
- `deploy/migrate_003_booking_ref.sql` adds `BookingRefID` (auto-numbered from 10001), shown as "Booking Ref: #10001" in emails.
- `deploy/migrate_004_booking_options.sql` adds `BookingOptions` (JSON list of the Admin Team's choices, including the Admin Team only fields) and `SelectedOption` (the one the manager approved).
- `deploy/migrate_005_travel_approver_mapping.sql` creates `TravelApproverMapping` (the portal's own routing table) and seeds it once from `checklist_empmanager_mapping`.
- `deploy/migrate_006_hotel_booking.sql` adds `TravelRequests.PreferredHotel` and creates `TravelBookingVouchers` (the uploaded PDF vouchers). **Required by this version.**

## Deploying under /dashboard/travel (behind the Mingalaba app on port 9000)

Port 9000 on `mingalaba.hungerpangs.com` is served by the Mingalaba app, so the portal cannot listen on it too. The portal runs on an internal port, and whatever serves port 9000 forwards `/dashboard/travel/` to it (a reverse proxy).

`.env` on the server:

```ini
APP_BASE_URL=http://mingalaba.hungerpangs.com:9000/dashboard/travel   # public URL incl. sub-path, no trailing slash
APP_PREFIX=/dashboard/travel                                          # must match the end of APP_BASE_URL
HOST=127.0.0.1                                                        # only the proxy on this server reaches the portal
PORT=8090                                                             # internal port (any free one)
TRUSTED_PROXY=127.0.0.1                                               # the proxy's address: trust its X-Forwarded-For
```

The proxy must forward the **full path** (do not strip `/dashboard/travel`) and send `X-Forwarded-For`. Examples:

**IIS** (URL Rewrite + Application Request Routing, ARR proxy enabled; ARR adds X-Forwarded-For by default), in the site that serves port 9000:

```xml
<rule name="Travel Portal" stopProcessing="true">
  <match url="^dashboard/travel(/.*)?$" />
  <action type="Rewrite" url="http://127.0.0.1:8090/dashboard/travel{R:1}" />
</rule>
```

**nginx**:

```nginx
location /dashboard/travel/ {
    proxy_pass http://127.0.0.1:8090;           # no path here: keeps /dashboard/travel/... intact
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
}
```

Check: `http://mingalaba.hungerpangs.com:9000/dashboard/travel/healthz` returns `ok`. User Management refuses requests where the proxy did not pass the visitor's IP.

## User Management (HR)

The **User Management** tab lets HR add, edit and delete rows in `dbo.TravelApproverMapping`, the portal's own table that decides who approves each employee's travel. Changes apply to new requests immediately.

`dbo.checklist_empmanager_mapping` belongs to the checklist apps: the portal only read it once, to seed its own table (migration 005), and never writes to it. Saves are refused if `MAPPING_TABLE` ever points at a `checklist_` table. After the seed the two tables are independent.

- Unlock with `HR_ADMIN_PASSWORD` from `.env`. The section locks again after `HR_SESSION_MINUTES` (30) of inactivity, or with **Lock**.
- 5 wrong passwords from one address pause sign-in for 15 minutes.
- By default it opens only from office/private networks (`HR_ALLOWED_NETWORKS`), even though the portal itself is public. Set it to `any` only once the portal runs on HTTPS.
- Leaving the manager code blank marks the employee as a HOD (no manager approval). Deleting a row has the same effect.
- Every add/edit/delete is written to `logs\portal.log` with the before/after values and the sender's IP.

**Export bookings** (same tab) downloads every hotel booking request as an Excel file: a full month, or custom From/To dates, filtered by request date or check-in date. It includes the status, approved hotel, cost, voucher date and the Admin Team only fields, so treat the file as confidential. Exports are logged.

## Security note (public IP over HTTP)

The portal is reachable from the internet on plain `http://`. Passwords (portal login, HR password) and approval links travel unencrypted, so anyone on the network path (public Wi-Fi, mobile carrier) can read them. Put HTTPS in front before wide use: e.g. a company domain with a certificate on IIS (the proxy in front of the portal), or ask IT for their preferred TLS setup. Then set `APP_BASE_URL` to the `https://` address.

## Deploy on 10.10.30.90

Prerequisites on the server: Python 3.10+ (installed for all users, on PATH) and *ODBC Driver 18 for SQL Server*.

1. Copy this folder to the server, e.g. `C:\Apps\TravelRequest` (git clone, or copy over RDP).
2. Put the `.env` file in that folder and replace every `<...>` placeholder.
3. In an **elevated** PowerShell, from that folder:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\deploy\install.ps1
   ```
   The script installs packages into `.venv`, checks the database connection, opens the `PORT` from `.env` in Windows Firewall (Domain/Private networks) unless `HOST=127.0.0.1` (behind the proxy, no rule is needed), and registers a `TravelPortal` scheduled task that starts the portal at boot and restarts it if it stops.

## Operate

| Task | Command (elevated PowerShell) |
|---|---|
| Status | `Get-ScheduledTask TravelPortal` |
| Restart (e.g. after editing `.env` or updating code) | `powershell -ExecutionPolicy Bypass -File .\deploy\restart.ps1` |
| Logs | `Get-Content .\logs\portal.log -Tail 50 -Wait` |
| Health check | `Invoke-WebRequest http://127.0.0.1:8090/dashboard/travel/healthz` |

To run it in the foreground for troubleshooting, stop the task (`Stop-ScheduledTask TravelPortal`, then end any `python.exe` still holding the portal's `PORT`) and run `.\.venv\Scripts\python app.py` instead.
