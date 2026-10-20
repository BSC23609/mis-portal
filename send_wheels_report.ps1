<#
  send_wheels_report.ps1  -  daily 10:00 AM Wheels India job-work stock statement.
  Queries WH45 (customer BSV01373) live from SAP (recipients/body read from
  the portal's Neon settings), renders a BSC-branded HTML table and emails it via
  Microsoft Graph as ai@bharatsteels.in (To/CC read from the same Neon settings).

  Schedule on the SAP box, weekdays 10:00. Run once by hand to test:
      powershell -ExecutionPolicy Bypass -File C:\scripts\send_wheels_report.ps1

  Reuses C:\scripts\config.ini for the SAP [database] block. Graph + Neon come from a
  new [wheels] section (see below). CRON_SECRET is NOT needed here.

      [wheels]
      neon        = postgresql://...neon connection string...
      tenant_id   = f3f819ba-724b-4c0b-a9c0-2aa8d12bcbcc
      client_id   = 627c4231-5cdf-40b8-8af8-fef565f62bb7
      client_secret = <the new Azure secret>
      sender      = pdqc@bharatsteels.in
      logo        = C:\scripts\BSC.png
      ; --- WhatsApp trigger (optional; leave out to skip) ---
      ; WATI creds + template name are set in Vercel env (WATI_API_URL, WATI_API_KEY,
      ; WHEELS_WA_TEMPLATE); the recipient numbers are set in the in-app Wheels Settings.
      ; Only the portal secret is needed here to fire the send:
      cron_secret = <same value as the CRON_SECRET env var on Vercel>
#>

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Cfg       = Join-Path $ScriptDir "config.ini"
$LogFile   = Join-Path $ScriptDir "wheels_mail.log"
function Log($m){ "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m" | Tee-Object -FilePath $LogFile -Append }

# ---- read config (sectioned) ----
$ini=@{}; $sec=$null
foreach($raw in Get-Content $Cfg -Encoding UTF8){ $l=$raw.Trim()
  if(-not $l -or $l[0] -in ';','#'){continue}
  if($l -match '^\[(.+)\]$'){ $sec=$matches[1].Trim(); $ini[$sec]=@{}; continue }
  if($sec -and $l -match '^([^=]+)=(.*)$'){ $ini[$sec][$matches[1].Trim()]=$matches[2].Trim() } }
$db=$ini['database']; $w=$ini['wheels']
if(-not $w){ Log "ERROR: [wheels] section missing in config.ini"; exit 1 }
$dbUser = if($db.username){$db.username}else{$db.user}

# ================= 0. skip Sundays and configured holidays =================
# No stock report goes out on a Sunday or on any date listed in the portal's Holidays setting.
$portalEarly = if($w.portal){ $w.portal } else { "https://mis.bharatsteels.in" }
$todayIso = (Get-Date).ToString('yyyy-MM-dd')
if((Get-Date).DayOfWeek -eq [DayOfWeek]::Sunday){ Log "Sunday ($todayIso) - no stock report. Exiting."; exit 0 }
try{
  $setEarly = Invoke-RestMethod -Uri "$($portalEarly.TrimEnd('/'))/api/wheels?route=settings" -TimeoutSec 30
  if(@($setEarly.holidays) -contains $todayIso){ Log "Holiday ($todayIso) - no stock report. Exiting."; exit 0 }
}catch{ Log "WARN: couldn't check holidays early ($($_.Exception.Message)) - continuing." }

# ================= 1. live WH45 stock for Wheels India =================
$connStr="Server=$($db.server);Database=$($db.database);User Id=$dbUser;Password=$($db.password);TrustServerCertificate=True;Connect Timeout=30"
$sql=@"
;WITH WIRItems AS (
    SELECT W.ItemCode, I.ItemName, I.U_Grade, I.U_Thick, I.U_Width, I.U_Length1, W.OnHand
    FROM OITW W INNER JOIN OITM I ON I.ItemCode=W.ItemCode
    WHERE W.WhsCode='45' AND W.OnHand>0
      AND (I.ItemName LIKE '%WIR0%' OR I.ItemName LIKE '%WIR %' OR I.ItemName LIKE '%WIR')
),
Batch AS (
    SELECT B.ItemCode, B.BatchNum, B.U_CoilNo, B.U_Mothercoil, B.InDate,
           ROW_NUMBER() OVER (PARTITION BY B.ItemCode ORDER BY B.Quantity DESC) rn,
           COUNT(*) OVER (PARTITION BY B.ItemCode) BatchCount
    FROM OIBT B WHERE B.WhsCode='45' AND B.Quantity<>0
)
SELECT WI.ItemName,
       ISNULL(BX.BatchNum,'') BatchNum, ISNULL(BX.U_CoilNo,'') CoilNo, ISNULL(BX.U_Mothercoil,'') Mother,
       ISNULL(WI.U_Grade,'') Grade, WI.U_Thick Thick, WI.U_Width Width, WI.U_Length1 Length,
       WI.OnHand Quantity, ISNULL(BX.BatchCount,0) BatchCount,
       DATEDIFF(day,BX.InDate,GETDATE()) AgeDays
FROM WIRItems WI
LEFT JOIN Batch BX ON BX.ItemCode=WI.ItemCode AND BX.rn=1
ORDER BY WI.ItemName
"@
$conn=New-Object System.Data.SqlClient.SqlConnection $connStr; $conn.Open()
$cmd=$conn.CreateCommand(); $cmd.CommandText=$sql; $cmd.CommandTimeout=120
$dt=New-Object System.Data.DataTable; [void](New-Object System.Data.SqlClient.SqlDataAdapter($cmd)).Fill($dt)
$conn.Close()
Log "SAP: $($dt.Rows.Count) WH45 batches for Wheels India"

# ================= 2. settings (recipients + body + enabled) from Neon =================
# psql via the Neon HTTP? Simplest on Windows: use the connection through .NET Npgsql if present,
# else call the portal API. We call the portal API (public settings route) to avoid a Postgres driver.
$portal = if($w.portal){ $w.portal } else { "https://mis.bharatsteels.in" }
try{
  $set = Invoke-RestMethod -Uri "$portal/api/wheels?route=settings" -TimeoutSec 30
}catch{ Log "ERROR: couldn't read settings from $portal/api/wheels: $($_.Exception.Message)"; exit 1 }

if($set.enabled -eq $false){ Log "Daily email disabled in settings - exiting."; exit 0 }
$to = @($set.mail_to); $cc = @($set.mail_cc)
$mailBody = [string]$set.mail_body; $mailSubject = [string]$set.mail_subject
if($to.Count -eq 0){ Log "No recipients configured - exiting."; exit 0 }
Log "Settings: To=$($to -join ',') CC=$($cc -join ',')"

# ================= 3. classify + build sections =================
function FormOf($n){ $n=$n.ToUpper()
  if($n -match 'SCRAP'){'SCRAP'} elseif($n -match 'BLANK|FLAT|STRIP'){'STRIP'}
  elseif($n -match 'PLATE'){'PLATE'} elseif($n -match 'COIL'){'COIL'} else {'OTHER'} }
$rows=@(); foreach($r in $dt.Rows){
  $f=(FormOf $r.ItemName); $q=[decimal]$r.Quantity
  if(($f -eq 'PLATE' -and $q -lt 0.1) -or ($f -eq 'STRIP' -and $q -lt 0.011)){ continue }   # plate <0.1 MT, strip <0.011 MT not listed
  $rows += [PSCustomObject]@{ Form=$f; Mill=($r.ItemName -split ' ')[0]; Name=$r.ItemName
    Grade=$r.Grade; T=$r.Thick; W=$r.Width; L=$r.Length; Coil=$r.CoilNo; Batch=$r.BatchNum
    Mother=$r.Mother; Qty=$q; Age=$r.AgeDays; BatchCount=$r.BatchCount } }

$sections=@(
  @{k='PLATE'; t='Plate Stock (WIP / FG)'; c='#0e7a4b'},
  @{k='STRIP'; t='Strip / Blank Stock';c='#b5651d'},
  @{k='COIL';  t='Coil Stock';         c='#1367a6'})   # WIP/FG (plate+strip) first, coil last; scrap excluded

function Dim($v){ if($v -is [DBNull] -or $null -eq $v -or "$v" -eq ''){return ''}; try{ ([double]$v).ToString('0.##') }catch{ '' } }
function Len($v){ if($v -is [DBNull] -or $null -eq $v -or "$v" -eq ''){return '-'}; try{ $n=[double]$v; if($n -gt 0){$n.ToString('0')}else{'-'} }catch{ '-' } }

$body=New-Object System.Text.StringBuilder
$gb=0; $gq=[decimal]0
foreach($s in $sections){
  $rs=@($rows | Where-Object Form -eq $s.k | Sort-Object Qty -Descending)
  if($rs.Count -eq 0){continue}
  $tq=($rs | Measure-Object Qty -Sum).Sum; $gb+=$rs.Count; $gq+=$tq
  [void]$body.Append("<tr><td colspan='9' style='background:$($s.c);color:#fff;font-weight:700;padding:7px 10px;font-size:13px'>$($s.t)<span style='float:right'>$($rs.Count) items &middot; $('{0:N3}' -f $tq) MT</span></td></tr>")
  [void]$body.Append("<tr style='background:#eef3f7'><th>#</th><th>Mill</th><th>Grade</th><th>Thk</th><th>Width</th><th>Length</th><th>Coil / Batch No</th><th>Qty (MT)</th><th>Age (d)</th></tr>")
  $i=0; foreach($x in $rs){ $i++
    $bc = if([int]("0"+[string]$x.BatchCount) -gt 1){" <span style='color:#999'>($($x.BatchCount) batches)</span>"}else{""}
    [void]$body.Append("<tr><td style='text-align:center;color:#888'>$i</td><td style='text-align:center'>$($x.Mill)</td><td style='text-align:center'>$($x.Grade)</td><td style='text-align:center'>$(Dim $x.T)</td><td style='text-align:center'>$(Dim $x.W)</td><td style='text-align:center'>$(Len $x.L)</td><td style='text-align:center;font-family:monospace;font-size:11px'>$(if($x.Coil){$x.Coil}else{$x.Batch})$bc</td><td style='text-align:center;font-weight:600'>$('{0:N3}' -f $x.Qty)</td><td style='text-align:center;color:#888'>$($x.Age)</td></tr>") }
  [void]$body.Append("<tr><td colspan='7' style='text-align:right;font-weight:700;padding:6px 10px;border-top:2px solid #333'>$($s.t) total</td><td style='text-align:center;font-weight:700;border-top:2px solid #333'>$('{0:N3}' -f $tq)</td><td style='border-top:2px solid #333'></td></tr>") }

$asof=(Get-Date).ToString('dd MMM yyyy, hh:mm tt')
$logoTag=""
if($w.logo -and (Test-Path $w.logo)){
  $b64=[Convert]::ToBase64String([IO.File]::ReadAllBytes($w.logo))
  $logoTag="<img src='data:image/png;base64,$b64' style='height:42px;background:#fff;padding:5px 8px;border-radius:4px'/>"
}
$html=@"
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>@media only screen and (max-width:600px){.wk-tbl{font-size:10px !important}.wk-tbl th,.wk-tbl td{padding:2px 3px !important}.wk-pad{padding:12px 8px 4px !important}}</style></head><body style="margin:0;padding:0;background:#fff;-webkit-text-size-adjust:100%">
<div style="font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;max-width:920px;width:100%;margin:auto;border:1px solid #d0d7dd;box-sizing:border-box">
  <div style="background:#0e2a47;color:#fff;padding:14px 22px;display:flex;align-items:center;gap:16px">
    $logoTag
    <div><div style="font-size:18px;font-weight:700">Job Work Stock Statement &mdash; Wheels India Limited</div>
    <div style="font-size:12px;opacity:.85;margin-top:3px">Bharat Steel (Chennai) Pvt. Ltd. &middot; Material held at BSC on customer's account</div></div>
  </div>
  $(if($mailBody.Trim()){ "<div style='padding:14px 22px 0;font-size:13px;color:#222;line-height:1.6'>$([System.Web.HttpUtility]::HtmlEncode($mailBody).Replace([Environment]::NewLine,'<br>').Replace("`n",'<br>'))</div>" })
  <div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;padding:11px 22px;background:#f7f9fb;border-bottom:1px solid #e3e8ec;font-size:13px">
    <div>Customer: <b>Wheels India Limited</b></div><div>As on: <b>$asof</b></div><div>Warehouse: <b>Job Work (WH 45)</b></div></div>
  <div class="wk-pad" style="padding:14px 22px 4px;overflow-x:auto;-webkit-overflow-scrolling:touch"><table class="wk-tbl" style="width:100%;border-collapse:collapse;font-size:12px">$($body.ToString())</table></div>
  <div style="margin:14px 22px;padding:12px 16px;background:#0e2a47;color:#fff;border-radius:5px;display:flex;justify-content:space-between;font-size:15px;font-weight:700"><span>TOTAL MATERIAL HELD</span><span>$gb items &middot; $('{0:N3}' -f $gq) MT</span></div>
  <div style="padding:6px 22px;font-size:9px;color:#aaa;font-style:italic">Computer generated Statement - may contain mistakes</div>
  <div style="padding:16px 22px;font-size:11px;color:#667;border-top:1px solid #e3e8ec;line-height:1.6"><b>Bharat Steel (Chennai) Pvt. Ltd.</b> &middot; No.147, Survey No.133, Thirunilai Village, Ponneri Taluk, Chennai 600103 &middot; 044 6791 7800 &middot; info@bharatsteels.in<br>System-generated statement of your material held at BSC as on the date/time shown. Figures drawn live from our ERP.</div>
</div></body></html>
"@


# ================= 3b. build Excel attachment (single Export-Excel call, robust) =================
if(-not (Get-Module -ListAvailable -Name ImportExcel)){
  try{ Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber } catch { Log "WARN: ImportExcel unavailable, sending without attachment: $($_.Exception.Message)" }
}
$xlsxPath = Join-Path $env:TEMP ("WheelsIndia_Stock_" + (Get-Date -Format 'yyyyMMdd_HHmm') + ".xlsx")
$attachB64 = $null; $attachName = Split-Path $xlsxPath -Leaf
if(Get-Module -ListAvailable -Name ImportExcel){
  Import-Module ImportExcel
  if(Test-Path $xlsxPath){ Remove-Item $xlsxPath -Force }
  function DVal($v){ if($v -is [DBNull] -or $null -eq $v -or "$v" -eq ''){return $null}; try{[double]$v}catch{$null} }
  $secTitle = [ordered]@{ PLATE='Plate Stock (WIP / FG)'; STRIP='Strip / Blank Stock'; COIL='Coil Stock' }
  $out = New-Object System.Collections.ArrayList
  $gGrand=[decimal]0; $bGrand=0
  foreach($sk in $secTitle.Keys){
    $rs = @($rows | Where-Object Form -eq $sk | Sort-Object Qty -Descending)
    if($rs.Count -eq 0){ continue }
    $tq = ($rs | Measure-Object Qty -Sum).Sum; $gGrand += $tq; $bGrand += $rs.Count
    $i=0
    foreach($x in $rs){ $i++
      [void]$out.Add([PSCustomObject][ordered]@{
        'Section' = $secTitle[$sk]
        'Mill'    = $x.Mill
        'Grade'   = $x.Grade
        'Thk'     = (DVal $x.T)
        'Width'   = (DVal $x.W)
        'Length'  = (DVal $x.L)
        'Coil / Batch No' = $(if($x.Coil){$x.Coil}else{$x.Batch})
        'Mother Coil'     = $(if($x.Mother){$x.Mother}else{'-'})
        'Qty (MT)'        = [math]::Round([double]$x.Qty,3)
        'Age (d)'         = $(if($x.Age -is [DBNull] -or "$($x.Age)" -eq ''){$null}else{[int]$x.Age})
      })
    }
    # section subtotal row
    [void]$out.Add([PSCustomObject][ordered]@{ 'Section'=$secTitle[$sk]; 'Mill'=''; 'Grade'=''; 'Thk'=$null; 'Width'=$null; 'Length'=$null; 'Coil / Batch No'=''; 'Mother Coil'=('Sub-total ('+$rs.Count+' items)'); 'Qty (MT)'=[math]::Round([double]$tq,3); 'Age (d)'=$null })
  }
  # grand total
  [void]$out.Add([PSCustomObject][ordered]@{ 'Section'='TOTAL'; 'Mill'=''; 'Grade'=''; 'Thk'=$null; 'Width'=$null; 'Length'=$null; 'Coil / Batch No'=''; 'Mother Coil'='TOTAL MATERIAL HELD (excl. scrap)'; 'Qty (MT)'=[math]::Round([double]$gGrand,3); 'Age (d)'=$null })

  try{
    $out | Export-Excel -Path $xlsxPath -WorksheetName "Stock Statement" -AutoSize -BoldTopRow -FreezeTopRow -TableName "Stock" -TableStyle Medium2
    $attachB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($xlsxPath))
    Log "Excel built: $attachName ($bGrand items, $('{0:N3}' -f $gGrand) MT)"
  }catch{ Log "WARN: Excel build failed, sending without attachment: $($_.Exception.Message)" }
}

# ================= 4. send via Microsoft Graph (app-only) =================
$tok = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($w.tenant_id)/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body @{
  client_id=$w.client_id; client_secret=$w.client_secret; scope="https://graph.microsoft.com/.default"; grant_type="client_credentials" }
$hdr=@{ Authorization="Bearer $($tok.access_token)"; "Content-Type"="application/json" }

$message = @{
    subject=$(if($mailSubject.Trim()){$mailSubject}else{"BSC - Wheels India Job Work Stock Statement - $((Get-Date).ToString('dd MMM yyyy'))"})
    body=@{ contentType="HTML"; content=$html }
    toRecipients=@($to | ForEach-Object { @{ emailAddress=@{ address=$_ } } })
    ccRecipients=@($cc | ForEach-Object { @{ emailAddress=@{ address=$_ } } })
  }
if($attachB64){
  $message.attachments = @(@{
    "@odata.type" = "#microsoft.graph.fileAttachment"
    name = $attachName
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    contentBytes = $attachB64
  })
}
$msg = @{ message=$message; saveToSentItems=$true } | ConvertTo-Json -Depth 10

try{
  Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($w.sender)/sendMail" -Headers $hdr -Body $msg
  Log "Sent. $gb items, $('{0:N3}' -f $gq) MT, to $($to.Count) + cc $($cc.Count)."
}catch{
  Log "ERROR sending: $($_.Exception.Message)"; exit 1
}

# ================= 4b. internal copy of the stock statement (Management + Production) =================
# Same Wheels India stock statement, sent to the internal lists (no customer CC).
# Each list gated by its own switch in settings. WhatsApp for these lists is handled by the portal trigger below.
$intTo = @()
if($set.mgmt_enabled -ne $false){ $intTo += @($set.mgmt_mail_to) }
if($set.prod_enabled -ne $false){ $intTo += @($set.prod_mail_to) }
$intTo = @($intTo | Where-Object { $_ } | Select-Object -Unique)
if($intTo.Count -gt 0){
  try{
    $intMessage = @{
      subject = "[Internal] " + $(if($mailSubject.Trim()){$mailSubject}else{"BSC - Wheels India Job Work Stock Statement - $((Get-Date).ToString('dd MMM yyyy'))"})
      body = @{ contentType="HTML"; content=$html }
      toRecipients = @($intTo | ForEach-Object { @{ emailAddress=@{ address=$_ } } })
    }
    if($attachB64){ $intMessage.attachments = @(@{ "@odata.type"="#microsoft.graph.fileAttachment"; name=$attachName; contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; contentBytes=$attachB64 }) }
    $intMsg = @{ message=$intMessage; saveToSentItems=$true } | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($w.sender)/sendMail" -Headers $hdr -Body $intMsg
    Log "Internal stock copy sent to $($intTo.Count) (mgmt+production)."
  }catch{ Log "Internal stock copy ERROR: $($_.Exception.Message)" }
}else{ Log "No internal stock recipients - internal copy skipped." }

# ================= 5. WhatsApp (triggered on the portal) =================
# The portal sends the same statement as a PDF via WATI. WATI credentials and the
# template name live in Vercel env (WATI_API_URL, WATI_API_KEY, WHEELS_WA_TEMPLATE);
# the recipient numbers are set in the in-app Wheels Settings page. This script just
# fires the trigger (guarded by CRON_SECRET). Needs in [wheels]: cron_secret.
if($w.cron_secret){
  try{
    $waResp = Invoke-RestMethod -Method POST -TimeoutSec 60 `
                -Uri "$($portal.TrimEnd('/'))/api/wheels" `
                -Headers @{ Authorization = "Bearer $($w.cron_secret)"; "Content-Type" = "application/json" } `
                -Body (@{ whatsapp = $true } | ConvertTo-Json)
    Log ("WhatsApp: " + ($waResp | ConvertTo-Json -Compress -Depth 6))
  }catch{ Log "WhatsApp trigger ERROR: $($_.Exception.Message)" }
}else{
  Log "WhatsApp not triggered ([wheels] cron_secret missing) - skipped."
}
