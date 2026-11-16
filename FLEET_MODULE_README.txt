FLEET MIS MODULE — files and how to deploy
==========================================

NEW FILES (drop into your mis-portal repo at these exact paths):
  api/fleet.js                 -> api/fleet.js         (backend, standalone)
  fleet/index.html             -> fleet/index.html     (dashboard + entry screens)
  fleet_schema.sql             -> run once in Neon SQL editor (creates tables + seeds vendor map)
  fleet_refresh.ps1            -> C:\scripts\fleet_refresh.ps1  (SAP cost engine; replaces the test one)

EDITED FILES (replace your copies with these — only small additions):
  _portal.js     (adds the "Fleet" nav tab + /fleet/ base path)
  vercel.json    (adds one rewrite so /api/fleet reaches api/fleet.js)

STEP 1 — database
  Open Neon SQL editor, paste fleet_schema.sql, run it once.

STEP 2 — deploy the app
  Copy the new/edited files into C:\github\mis-portal, then run push.bat.
  (Vercel auto-deploys. The "Fleet" tab appears in the top nav.)

STEP 3 — SAP cost feed
  Put fleet_refresh.ps1 in C:\scripts. In C:\scripts\config.ini add a [portal] section:
     [portal]
     base=https://mis.bharatsteels.in
     cron_secret=<same CRON_SECRET as in Vercel>
  Run once:  powershell -ExecutionPolicy Bypass -File C:\scripts\fleet_refresh.ps1
  It fills Diesel/Bata/Maintenance per vehicle into the dashboard.
  Schedule it (Task Scheduler) hourly or daily like the stock refresh.

STEP 4 — enter the rest in the app
  Vehicle Master tab : driver + salary (monthly) + insurance/road-tax/fc/depreciation/interest (yearly).
  Daily Trips tab    : one row per trip (income blank for empty trips).
  Monthly Toll tab   : toll per vehicle per month.
  Dashboard tab      : the monthly P&L matrix (pick a month).

NOTES
  - Insurance booked inside the SAP vehicle accounts is EXCLUDED from the variable split
    (you enter it as Fixed in Vehicle Master) — no double counting.
  - "Unclassified" SAP lines are shown under the dashboard for you to review.
  - Vendor map is seeded in fleet_schema.sql (edit the table if a vendor changes category).
