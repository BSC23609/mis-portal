<#
  send_production_report.ps1  -  daily INTERNAL Wheels India production report.
  Separate from send_wheels_report.ps1 (that one is the CUSTOMER stock statement).

  This script is internal-only: it emails the production report (image or PDF) to the
  internal recipients set in the portal, and fires the production WhatsApp trigger.
  The portal renders the report; this script just mails it and pokes the WhatsApp route.

  Schedule on the SAP box as its own task (e.g. weekdays 10:05). Run once by hand to test:
      powershell -NoProfile -ExecutionPolicy Bypass -File C:\scripts\send_production_report.ps1

  Reuses C:\scripts\config.ini  [wheels]  (tenant_id, client_id, client_secret, sender,
  portal, cron_secret). No SAP DB connection needed here.
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Cfg       = Join-Path $ScriptDir "config.ini"
$LogFile   = Join-Path $ScriptDir "production_mail.log"
function Log($m){ "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m" | Tee-Object -FilePath $LogFile -Append }

# ---- read config (sectioned) ----
$ini=@{}; $sec=$null
foreach($raw in Get-Content $Cfg -Encoding UTF8){ $l=$raw.Trim()
  if(-not $l -or $l[0] -in ';','#'){continue}
  if($l -match '^\[(.+)\]$'){ $sec=$matches[1].Trim(); $ini[$sec]=@{}; continue }
  if($sec -and $l -match '^([^=]+)=(.*)$'){ $ini[$sec][$matches[1].Trim()]=$matches[2].Trim() } }
$w=$ini['wheels']
if(-not $w){ Log "ERROR: [wheels] section missing in config.ini"; exit 1 }
$portal = if($w.portal){ $w.portal } else { "https://mis.bharatsteels.in" }

# ---- read production settings from the portal ----
try{
  $set = Invoke-RestMethod -Uri "$($portal.TrimEnd('/'))/api/wheels?route=settings" -TimeoutSec 30
}catch{ Log "ERROR: couldn't read settings from $portal: $($_.Exception.Message)"; exit 1 }

# ---- skip Sundays and configured holidays (no report on those days) ----
$todayIso = (Get-Date).ToString('yyyy-MM-dd')
if((Get-Date).DayOfWeek -eq [DayOfWeek]::Sunday){ Log "Sunday ($todayIso) - no production report. Exiting."; exit 0 }
if(@($set.holidays) -contains $todayIso){ Log "Holiday ($todayIso) - no production report. Exiting."; exit 0 }

# Production Report recipients = Management (if enabled) + Production (if enabled), de-duplicated.
$prodTo = @()
if($set.mgmt_enabled -ne $false){ $prodTo += @($set.mgmt_mail_to) }
if($set.prod_enabled -ne $false){ $prodTo += @($set.prod_mail_to) }
$prodTo = @($prodTo | Where-Object { $_ } | Select-Object -Unique)
$prodFmt = if($set.prod_format){ [string]$set.prod_format } else { 'image' }
Log "Production: format=$prodFmt, mail_to=$($prodTo.Count) (mgmt+production)"

# ================= 1. internal production email (MS Graph) =================
if($prodTo.Count -gt 0){
  try{
    $tok = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($w.tenant_id)/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body @{
      client_id=$w.client_id; client_secret=$w.client_secret; scope="https://graph.microsoft.com/.default"; grant_type="client_credentials" }
    $hdr=@{ Authorization="Bearer $($tok.access_token)"; "Content-Type"="application/json" }

    $pExt = if($prodFmt -eq 'pdf'){ '.pdf' } else { '.png' }
    $pUrl = "$($portal.TrimEnd('/'))/wheels-production$pExt"
    $pTmp = Join-Path $env:TEMP ('wheels_prod_'+[Guid]::NewGuid().ToString('N')+$pExt)
    Invoke-WebRequest -Uri $pUrl -OutFile $pTmp -TimeoutSec 60 -UseBasicParsing
    $pB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($pTmp))
    Remove-Item $pTmp -ErrorAction SilentlyContinue
    $stamp = (Get-Date).ToString('yyyyMMdd')
    if($prodFmt -eq 'pdf'){
      $pName='WheelsIndia_Production_'+$stamp+'.pdf'; $pType='application/pdf'
      $pHtml='<p>Dear Team,</p><p>Please find attached the Wheels India Production Report (PDF).</p><p>Bharat Steel (Chennai) Pvt. Ltd.</p>'
      $pAtt=@(@{ '@odata.type'='#microsoft.graph.fileAttachment'; name=$pName; contentType=$pType; contentBytes=$pB64 })
    }else{
      $pName='WheelsIndia_Production_'+$stamp+'.png'; $pType='image/png'
      $pHtml='<p>Dear Team,</p><p>Wheels India Production Report:</p><img src="cid:prodimg" style="max-width:100%;border:1px solid #d0d7dd"/><p>Bharat Steel (Chennai) Pvt. Ltd.</p>'
      $pAtt=@(@{ '@odata.type'='#microsoft.graph.fileAttachment'; name=$pName; contentType=$pType; contentBytes=$pB64; contentId='prodimg'; isInline=$true })
    }
    $pMessage=@{
      subject='Wheels India - Production Report - '+(Get-Date).ToString('dd MMM yyyy')
      body=@{ contentType='HTML'; content=$pHtml }
      toRecipients=@($prodTo | ForEach-Object { @{ emailAddress=@{ address=$_ } } })
      attachments=$pAtt
    }
    $pMsg=@{ message=$pMessage; saveToSentItems=$true } | ConvertTo-Json -Depth 12
    Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($w.sender)/sendMail" -Headers $hdr -Body $pMsg
    Log "Production email sent ($prodFmt) to $($prodTo.Count)."
  }catch{ Log "Production email ERROR: $($_.Exception.Message)" }
}else{ Log "No production email recipients - email skipped." }

# ================= 2. internal production WhatsApp (portal trigger) =================
if($w.cron_secret){
  try{
    $pwaResp = Invoke-RestMethod -Method POST -TimeoutSec 60 `
                 -Uri "$($portal.TrimEnd('/'))/api/wheels" `
                 -Headers @{ Authorization = "Bearer $($w.cron_secret)"; "Content-Type" = "application/json" } `
                 -Body (@{ prod_whatsapp = $true } | ConvertTo-Json)
    Log ("Production WhatsApp: " + ($pwaResp | ConvertTo-Json -Compress -Depth 6))
  }catch{ Log "Production WhatsApp trigger ERROR: $($_.Exception.Message)" }
}else{
  Log "Production WhatsApp not triggered ([wheels] cron_secret missing) - skipped."
}
