# Deploying the Travel Accommodation Portal on 10.10.30.90

Final address: **http://mingalaba.hungerpangs.com:9000/dashboard/travel/**

IIS already serves port 9000 (the Mingalaba app). The portal runs on the same server at `127.0.0.1:8090`, and an IIS rule forwards `/dashboard/travel/...` to it. Mingalaba is not changed.

Time needed: about 20 minutes. You need to sign in to the server as an administrator (Remote Desktop).

---

## Before you start (check once)

On 10.10.30.90:

| Needed | How to check | If missing |
|---|---|---|
| Python 3.10 or newer, installed **for all users** and on PATH | `python --version` in PowerShell | Install from python.org: tick **"Add python.exe to PATH"**, choose **Customize installation > Install for all users** |
| ODBC Driver 18 for SQL Server | `Get-OdbcDriver -Name "ODBC Driver 18 for SQL Server"` | Download from Microsoft ("ODBC Driver 18 for SQL Server") and install |
| IIS URL Rewrite + Application Request Routing (ARR) | Already used by Mingalaba on port 9000 | Nothing to do |

Database: migrations 001 to 005 are **already applied** to MingalabaDB. **Migration 006** (`deploy/migrate_006_hotel_booking.sql`: the Preferred Hotel column and the voucher table) must be applied before this version runs; it is safe to re-run. Open it in SQL Server Management Studio connected to MingalabaDB and click **Execute**.

---

## Step 1: Copy the files

1. Copy `TravelPortal_deploy.zip` to the server over Remote Desktop, a USB drive or a shared folder.
   **Do not email it:** it contains `.env` with passwords and keys.
2. Right-click the zip, choose **Extract All**, and extract to `C:\Apps\TravelRequest`.
   The folder should contain `app.py`, `.env`, `templates`, `static` and `deploy`.

## Step 2: Check `.env`

Open `C:\Apps\TravelRequest\.env` in Notepad.

1. Change `PORTAL_PASSWORD` to a long password (12+ characters). Everyone uses it to open the portal, and the portal is reachable from the internet.
2. Check these lines read exactly:

```ini
APP_BASE_URL=http://mingalaba.hungerpangs.com:9000/dashboard/travel
APP_PREFIX=/dashboard/travel
HOST=127.0.0.1
PORT=8090
TRUSTED_PROXY=127.0.0.1
DB_SERVER=127.0.0.1
MAPPING_TABLE=dbo.TravelApproverMapping
```

3. Leave `SECRET_KEY` unchanged: changing it breaks every approval link already emailed.
4. Save.

## Step 3: Install and start the portal

1. Click Start, type **PowerShell**, right-click it, and choose **Run as administrator**.
2. Run:

```powershell
cd C:\Apps\TravelRequest
powershell -ExecutionPolicy Bypass -File .\deploy\install.ps1
```

This creates `.venv`, installs packages, checks the database, and registers a **TravelPortal** scheduled task that starts the portal now and at every reboot.

It should end with a green line: `Portal is running (health: ok)`.
If it stops with an error, the message says what is missing (Python, ODBC driver, `.env` value, database).

## Step 4: Add the IIS rule

In the same administrator PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\deploy\add_iis_rule.ps1
```

The script:
- finds the IIS site on port 9000
- backs up the IIS configuration and the site's `web.config`
- adds a rule named **Travel Portal** as the **first** rule, so it runs before the Mingalaba rule (otherwise Mingalaba redirects to its login)
- checks the result through IIS

It should end with: `Check through IIS: http://localhost:9000/dashboard/travel/healthz -> ok`.

If it warns that **X-Forwarded-For** is not sent, ask IT to turn on *IIS Manager > (server) > Application Request Routing Cache > Server Proxy Settings > Preserve client IP in the following header: X-Forwarded-For*. This is a server-wide setting, so check it with IT first. Without it, User Management refuses everyone; the rest of the portal works.

## Step 5: Test

From any PC or phone:

| Open | Expect |
|---|---|
| http://mingalaba.hungerpangs.com:9000/dashboard/travel/healthz | `ok` |
| http://mingalaba.hungerpangs.com:9000/dashboard/travel/ | Portal login prompt, then **Hotel booking request** |
| http://mingalaba.hungerpangs.com:9000/ | The Mingalaba login, as before |

Then do one full test request:
1. Submit a request with your own Emp Code.
2. The Admin Team email arrives. Click **Enter Booking Details**, fill in a choice (including the Admin Team only fields), and click **Save & Send for Approval**.
3. The manager email arrives. Click **Approve**, tick the non-refundable box and click **Confirm Approval**.
4. The Admin Team gets "Approved - final booking needed". Click **Upload Booking Voucher** and upload a PDF.
5. The employee receives the voucher as an attachment, with the manager in CC.

From an **office PC**, also open the **User Management** tab, unlock it with `HR_ADMIN_PASSWORD` from `.env`, and try **Export bookings**. From outside the office it shows "Office network only"; that is expected.

---

## Day-to-day

Run these in an administrator PowerShell in `C:\Apps\TravelRequest`.

| Task | Command |
|---|---|
| Restart (after editing `.env` or updating files) | `powershell -ExecutionPolicy Bypass -File .\deploy\restart.ps1` |
| See the log | `Get-Content .\logs\portal.log -Tail 50 -Wait` |
| Is it running? | `Get-ScheduledTask TravelPortal` |

**Updating the portal later:** stop it (`Stop-ScheduledTask TravelPortal`), copy the new files over the folder (keep your `.env`), then run `.\deploy\install.ps1` again.

## Undo (if something goes wrong)

- **Remove the IIS rule:** IIS Manager > the port 9000 site (or the server, if the script said "global rules") > **URL Rewrite** > select **Travel Portal** > **Remove**. Mingalaba is unaffected either way.
  Or restore the backup the script made: `Restore-WebConfiguration -Name TravelPortal-before-rule-<date-time>`.
- **Stop the portal:** `Stop-ScheduledTask TravelPortal` and `Unregister-ScheduledTask TravelPortal -Confirm:$false`.

## Security reminder

The site runs on plain `http://`, so passwords and approval links travel unencrypted over the internet. Ask IT to add an HTTPS certificate to the port 9000 site. After that, change `APP_BASE_URL` to `https://...` and restart the portal.
