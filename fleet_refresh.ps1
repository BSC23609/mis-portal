<#
  fleet_refresh.ps1 - READ-ONLY SAP pull + classifier for the Fleet MIS module.
  Pulls every posting to the per-vehicle COMMERCIAL VEHICLE EXP accounts and splits it
  into Diesel / Driver Bata / Maintenance per vehicle per month, using the confirmed
  vendor map (plus memo keywords for card/driver lines). Insurance vendors are excluded
  (those are entered as Fixed cost in Vehicle Master). Unclassified lines are reported for
  review so nothing is ever silently dropped. Changes nothing in SAP.

  Edit $Start / $End if needed, then:
     powershell -ExecutionPolicy Bypass -File C:\scripts\fleet_refresh.ps1
  Outputs C:\scripts\fleet_costs.csv  and prints an April check vs your manual sheet.
#>
$ErrorActionPreference="Stop"
$Cfg="C:\scripts\config.ini"; $Out="C:\scripts\fleet_refresh.txt"; $Csv="C:\scripts\fleet_costs.csv"
$Start="2026-04-01"      # first month to pull (inclusive)
$End  ="2026-12-01"      # exclusive upper bound
if(Test-Path $Out){ Remove-Item $Out -Force }

$ini=@{}; Get-Content $Cfg | ForEach-Object { if($_ -match '^\s*([^=;#\[]+)=(.*)$'){ $ini[$matches[1].Trim()]=$matches[2].Trim() } }
$dbUser=if($ini.username){$ini.username}else{$ini.user}
if(-not $ini.server -or -not $dbUser){ Write-Host "ERROR: config.ini missing server/username"; exit 1 }
$connStr="Server=$($ini.server);Database=$($ini.database);User Id=$dbUser;Password=$($ini.password);TrustServerCertificate=True;Connect Timeout=30"
function Invoke-Sql($sql){ $c=New-Object System.Data.SqlClient.SqlConnection $connStr; $c.Open()
  $cmd=$c.CreateCommand(); $cmd.CommandText=$sql; $cmd.CommandTimeout=300
  $da=New-Object System.Data.SqlClient.SqlDataAdapter $cmd; $dt=New-Object System.Data.DataTable; [void]$da.Fill($dt); $c.Close(); return ,$dt }

# ---- confirmed vendor map (CardCode -> category) ----
$sql = @"
DECLARE @Start DATE='$Start'; DECLARE @End DATE='$End';
;WITH VExp AS (
  SELECT
    LTRIM(REPLACE(REPLACE(a.AcctName,'COMMERCIAL VEHICLE EXP',''),' ','')) AS Vehicle,
    DATEFROMPARTS(YEAR(t.RefDate),MONTH(t.RefDate),1) AS Mon,
    CAST(t.Debit - t.Credit AS DECIMAL(18,2)) AS Amt,
    p.CardCode AS Vendor,
    ISNULL(t.LineMemo,'') AS Memo
  FROM JDT1 t
  JOIN OACT a ON a.AcctCode = t.Account
  LEFT JOIN OPCH p ON p.TransId = t.TransId
  WHERE a.AcctName LIKE 'COMMERCIAL VEHICLE EXP%'
    AND t.RefDate >= @Start AND t.RefDate < @End
),
Cls AS (
  SELECT Vehicle, Mon, Amt, Vendor, Memo,
    CASE
      WHEN Vendor IN ('BSV00843','BSV00773') THEN 'Insurance'
      WHEN Vendor IN ('BSV00190','BSV00133','BSV01312','BSV01409','BSV00761','BSV00763') THEN 'Maintenance'
      WHEN Vendor IN ('BSV00220','BSV01187','BSV01253','BSV00784','BSV01378','BSV00221') THEN 'Bata'
      WHEN Memo LIKE '%DRIVER%' OR Memo LIKE '%BATA%' OR Memo LIKE '%LOAD%' OR Memo LIKE '%UNLOAD%' THEN 'Bata'
      WHEN Memo LIKE '%CARD%' OR Memo LIKE '%DIESEL%' OR Memo LIKE '%FUEL%' OR Memo LIKE '%PETROL%' THEN 'Diesel'
      WHEN Memo LIKE '%REPAIR%' OR Memo LIKE '%MAINT%' OR Memo LIKE '%TYRE%' OR Memo LIKE '%SERVICE%' OR Memo LIKE '%SPARE%' OR Memo LIKE '%BATTERY%' THEN 'Maintenance'
      ELSE 'Unclassified'
    END AS Cat
  FROM VExp
)
SELECT Vehicle, Mon,
  CAST(SUM(CASE WHEN Cat='Diesel'       THEN Amt ELSE 0 END) AS DECIMAL(18,2)) AS Diesel,
  CAST(SUM(CASE WHEN Cat='Bata'         THEN Amt ELSE 0 END) AS DECIMAL(18,2)) AS Bata,
  CAST(SUM(CASE WHEN Cat='Maintenance'  THEN Amt ELSE 0 END) AS DECIMAL(18,2)) AS Maintenance,
  CAST(SUM(CASE WHEN Cat='Unclassified' THEN Amt ELSE 0 END) AS DECIMAL(18,2)) AS Unclassified
FROM Cls
WHERE Cat <> 'Insurance'
GROUP BY Vehicle, Mon
ORDER BY Mon, Vehicle;
"@

("##### Fleet cost engine - "+(Get-Date).ToString("yyyy-MM-dd HH:mm")+" #####")|Tee-Object -FilePath $Out -Append
("Window: "+$Start+" to "+$End)|Tee-Object -FilePath $Out -Append
try{
  $dt = Invoke-Sql $sql
  ($dt | Format-Table -AutoSize | Out-String -Width 400).TrimEnd() | Tee-Object -FilePath $Out -Append
  $dt | Export-Csv -Path $Csv -NoTypeInformation -Encoding UTF8
  ("CSV written: "+$Csv)|Tee-Object -FilePath $Out -Append

  # ---- push to the Fleet MIS module (Neon via /api/fleet?fmt=sync) ----
  if($ini.base -and $ini.cron_secret){
    $payload=@{ rows=@() }
    foreach($r in $dt.Rows){
      $payload.rows += @{ month=([datetime]$r.Mon).ToString('yyyy-MM-01'); vehicle=[string]$r.Vehicle;
        diesel=[double]$r.Diesel; bata=[double]$r.Bata; maintenance=[double]$r.Maintenance; unclassified=[double]$r.Unclassified }
    }
    $body=$payload|ConvertTo-Json -Depth 5 -Compress
    $uri=($ini.base.TrimEnd('/'))+'/api/fleet?fmt=sync'
    try{
      $resp=Invoke-RestMethod -Uri $uri -Method Post -Headers @{ Authorization='Bearer '+$ini.cron_secret } -ContentType 'application/json' -Body $body
      ("Pushed to portal: "+($resp|ConvertTo-Json -Compress))|Tee-Object -FilePath $Out -Append
    }catch{ ("Portal push FAILED: "+$_.Exception.Message)|Tee-Object -FilePath $Out -Append }
  } else {
    "config.ini [portal] base/cron_secret not set - CSV written, not pushed to portal."|Tee-Object -FilePath $Out -Append
  }

  # ---- April check vs your manual VEHICLE MONTHLY MIS sheet ----
  ("`r`n--- APRIL 2026 check (SAP classifier vs your manual sheet) ---")|Tee-Object -FilePath $Out -Append
  $manual=@{
    "TN04AE3484"=@(12565,16100,0); "TN04AF6394"=@(22556,20500,0); "TN04AH1119"=@(36968,31100,0);
    "TN18BM8443"=@(36968,24600,0); "TN18BV5968"=@(44563,25950,0); "TN18CA0415"=@(25044,22440,0); "TN04AD3336"=@(0,0,94352)
  }
  ("{0,-12} {1,12} {2,12} {3,12} {4,10} | manual D/B/M" -f "Vehicle","Diesel","Bata","Maint","Unclass")|Tee-Object -FilePath $Out -Append
  foreach($r in $dt.Rows){
    if([string]$r.Mon -like "*04-01*" -or ([datetime]$r.Mon).Month -eq 4){
      $v=[string]$r.Vehicle; $m=if($manual.ContainsKey($v)){$manual[$v]}else{@("-","-","-")}
      ("{0,-12} {1,12:N0} {2,12:N0} {3,12:N0} {4,10:N0} | {5}/{6}/{7}" -f $v,$r.Diesel,$r.Bata,$r.Maintenance,$r.Unclassified,$m[0],$m[1],$m[2])|Tee-Object -FilePath $Out -Append
    }
  }
}catch{ ("ERROR: "+$_.Exception.Message)|Tee-Object -FilePath $Out -Append }
Write-Host "`r`nDone. Paste back $Out"
