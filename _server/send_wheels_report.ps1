# send_wheels_report.ps1  -  Morning 10:00 Job Work auto-send (pinger).
#
# Hits the portal's /api/am-cron, which sends — to the recipients configured on the
# Settings page — the Job Work Stock, Daily Stock and Production reports, each on
# BOTH email and WhatsApp. The portal does all the work; this script just triggers it.
#
# Safe to run once at 10:00 or repeatedly: the endpoint is idempotent (each report+channel
# sends at most once per working day), time-guarded (does nothing before 10:00 IST) and
# skips Sundays + holidays. Guarded by cron_secret, same as jw_cron.ps1.
#
# Scheduled task: "BSC Wheels Daily Send", weekdays 10:00, S4U / run whether logged on or not.
# (The canonical copy of this file lives in the mis-portal repo under _server/.)

$c=@{}
Get-Content C:\scripts\config.ini | ForEach-Object {
  if($_ -match '^\s*([^=;#\[]+)=(.*)$'){ $c[$matches[1].Trim()] = $matches[2].Trim() }
}
$base = if($c.portal){$c.portal}elseif($c.base){$c.base}else{'https://mis.bharatsteels.in'}
Invoke-RestMethod -Uri ($base.TrimEnd('/') + '/api/am-cron') -Headers @{ Authorization = 'Bearer ' + $c.cron_secret } -TimeoutSec 180
