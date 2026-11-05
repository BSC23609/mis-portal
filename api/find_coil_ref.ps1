<#
  find_coil_ref.ps1  -  lists the candidate "coil reference" columns on the
  Wheels India production-schedule tables, so we can add the right one to the
  Production Plan. Read-only - selects nothing but column names + a few samples.

  Run:
      powershell -NoProfile -ExecutionPolicy Bypass -File C:\scripts\find_coil_ref.ps1
  Copy ALL output back.
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Cfg = Join-Path $ScriptDir "config.ini"
$ini=@{}; Get-Content $Cfg | ForEach-Object { if($_ -match '^\s*([^=;#\[]+)=(.*)$'){ $ini[$matches[1].Trim()]=$matches[2].Trim() } }
$dbUser = if($ini.username){$ini.username}else{$ini.user}
$connStr="Server=$($ini.server);Database=$($ini.database);User Id=$dbUser;Password=$($ini.password);TrustServerCertificate=True;Connect Timeout=30"

function Run($sql){
  $conn=New-Object System.Data.SqlClient.SqlConnection $connStr; $conn.Open()
  $cmd=$conn.CreateCommand(); $cmd.CommandText=$sql; $cmd.CommandTimeout=120
  $da=New-Object System.Data.SqlClient.SqlDataAdapter $cmd
  $dt=New-Object System.Data.DataTable; [void]$da.Fill($dt); $conn.Close()
  return $dt
}

Write-Host "=== 1. Columns that look like a coil / batch / reference on the schedule tables ==="
$sql1 = @'
SELECT t.name AS TableName, c.name AS ColumnName, ty.name AS DataType
FROM sys.columns c
JOIN sys.tables t  ON t.object_id = c.object_id
JOIN sys.types  ty ON ty.user_type_id = c.user_type_id
WHERE t.name IN ('@SMS_OSCH','@SMS_SCH1','@SMS_OSHO','@SMS_SHO1')
  AND ( c.name LIKE '%Coil%' OR c.name LIKE '%Batch%' OR c.name LIKE '%Ref%'
     OR c.name LIKE '%Heat%' OR c.name LIKE '%Mother%' OR c.name LIKE '%Roll%' OR c.name LIKE '%Lot%' )
ORDER BY t.name, c.name
'@
Run $sql1 | Format-Table -AutoSize | Out-String -Width 400 | Write-Host

Write-Host ""
Write-Host "=== 2. ALL user (U_) columns on the input-line and master tables (in case it is named differently) ==="
$sql2 = @'
SELECT t.name AS TableName, c.name AS ColumnName, ty.name AS DataType
FROM sys.columns c
JOIN sys.tables t  ON t.object_id = c.object_id
JOIN sys.types  ty ON ty.user_type_id = c.user_type_id
WHERE t.name IN ('@SMS_SCH1','@SMS_OSCH','@SMS_SHO1')
  AND c.name LIKE 'U[_]%'
ORDER BY t.name, c.name
'@
Run $sql2 | Format-Table -AutoSize | Out-String -Width 400 | Write-Host

Write-Host ""
Write-Host "=== 3. Sample: input line (U_ fields) of 5 recent OPEN schedules - find the one holding the coil no ==="
$sql3 = @'
SELECT TOP 5
    H.DocNum AS Master,
    OH.DocNum AS OutSch,
    LEFT(S.U_IPDesrp, 35) AS InputItem,
    S.U_IPItem           AS InputCode,
    S.U_IOnStock         AS InputQty
FROM [@SMS_OSHO] OH
LEFT JOIN [@SMS_SCH1] S ON TRY_CONVERT(INT, S.U_SHONum) = OH.DocNum
LEFT JOIN [@SMS_OSCH] H ON H.DocEntry = S.DocEntry
WHERE OH.Status='O'
ORDER BY OH.DocNum DESC
'@
Run $sql3 | Format-Table -AutoSize | Out-String -Width 400 | Write-Host

Write-Host "=== Done. Paste all output back - I will add the coil reference column to the plan. ==="
