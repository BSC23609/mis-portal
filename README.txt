BSC mis-portal update
=====================

A) PORTAL FILES  ->  copy into  C:\github\mis-portal  (keep the folders), then run push.bat
   - api\index.js            (Production Plan: coil column, auto-cron 6:30pm, honest WhatsApp log,
                              email without pending list, plan recipients in Settings)
   - vercel.json             (/wheels-plan.png + /api/plan-cron routes, 6:30pm Vercel cron)
   - wheels\plan.html        (Production Plan tab - Input coil column)
   - wheels\settings.html    (Production Plan recipients + test buttons)
   - wheels\log.html         (Send Log shows the Production Plan card)
   - stock\monthly.html      (Oldest coils - priority to move panel)

B) SERVER SCRIPT  ->  copy to  C:\scripts  on the SAP box, run once by hand
   - refresh_stock.ps1       (adds Coil No / Coil Item to wheels_prod_schedule.csv)
     Run it FIRST so the CSV has the coil columns, THEN push the portal files.

NOTES
 - The Production Plan now sends itself at 6:30pm via the Vercel cron in vercel.json.
   You do NOT need the old send_plan_report.ps1 / setup_plan_task.ps1 (Task Scheduler) anymore.
 - Vercel env still needed: CRON_SECRET, WHEELS_WA_PLAN_TEMPLATE (plan WhatsApp template),
   and at least one recipient ticked for Production Plan in Settings.
 - Daily Stock not delivering is a WATI matter: template bsc_daily_stock_report is Marketing;
   recreate it as Utility and repoint WHEELS_WA_DAILY_TEMPLATE. (No file in this zip.)
