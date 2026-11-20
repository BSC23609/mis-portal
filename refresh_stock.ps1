<#
  refresh_stock.ps1 - stock / sales / purchase pipeline for stock.bharatsteels.in
  Queries SAP MSSQL, writes coils.csv + plates.csv (HR stock pages) and
  mis_stock.csv + mis_sales.csv + mis_purchase.csv + mis_meta.json (MIS page)
  into the repo, then pushes to GitHub. One job, one push.
  Run from PowerShell (NOT cmd - a .ps1 opens in Notepad from cmd):
      powershell -ExecutionPolicy Bypass -File C:\scripts\refresh_stock.ps1
  Schedule it the same way as the weighbridge refresh.

  Config: reuse C:\scripts\config.ini  (gitignored) with:
      [database]
      server=10.10.250.11
      database=BSC_FINAL_2707
      username=dbadmin
      password=YOURPASS
#>

$ErrorActionPreference = "Stop"
# Never let git block on a credential prompt in the background (a prompt would hang the whole
# run forever and freeze every feed). With these set, git fails fast instead of waiting for input.
$env:GIT_TERMINAL_PROMPT = "0"
$env:GCM_INTERACTIVE     = "never"
$env:GIT_ASKPASS         = ""
$RepoDir  = "C:\github\bsc-stock-data"
$StockDir = $RepoDir
$Cfg      = "C:\scripts\config.ini"
$LogFile  = "C:\scripts\refresh_stock.log"

# ---- dispatch (weighbridge) feed -> mis-portal repo -------------------------
# This merges the old, flaky standalone weighbridge job into this single 15-min
# job: one SAP connection, both feeds, one schedule. The dispatch CSV lives in a
# DIFFERENT repo (mis-portal, which serves mis.bharatsteels.in), so it needs its
# own clone. Point $DispatchRepo at your existing mis-portal clone; if the folder
# is empty/absent the script clones it once from $DispatchRepoUrl.
$DispatchRepo    = "C:\github\mis-portal"
$DispatchRepoUrl = "https://github.com/BSC23609/mis-portal.git"

function Log($m){ "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m" | Tee-Object -FilePath $LogFile -Append }

# ---- read config.ini ----
$ini=@{}; Get-Content $Cfg | ForEach-Object {
  if($_ -match '^\s*([^=;#\[]+)=(.*)$'){ $ini[$matches[1].Trim()]=$matches[2].Trim() }
}
# accept either  user=  (old header)  or  username=  (weighbridge / discovery config)
$dbUser = if($ini.username){$ini.username}else{$ini.user}
if(-not $ini.server -or -not $dbUser){ Log "ERROR: config.ini missing server/username - check $Cfg"; exit 1 }
$connStr="Server=$($ini.server);Database=$($ini.database);User Id=$dbUser;Password=$($ini.password);TrustServerCertificate=True;Connect Timeout=30"

function Invoke-Sql($sql){
  $conn=New-Object System.Data.SqlClient.SqlConnection $connStr
  $conn.Open()
  $cmd=$conn.CreateCommand(); $cmd.CommandText=$sql; $cmd.CommandTimeout=120
  $da=New-Object System.Data.SqlClient.SqlDataAdapter $cmd
  $dt=New-Object System.Data.DataTable; [void]$da.Fill($dt)
  $conn.Close(); return ,$dt   # comma stops PowerShell unrolling the table into rows
}

# ---- brand parsed from item name (mill), not the steel grade UDF ----
function Get-Brand($name){
  $n=($name -replace 'H\.R','HR').ToUpper().Trim()
  foreach($b in @('NMDC','SAIL','JSW','AMNS','RINL','JINDAL','TATA','POSCO')){
    if($n.StartsWith($b)){ return $b }
  }
  return 'OTHER'
}

# ======================= QUERIES =======================
$coilSql = @"
SELECT
    T0.ItemCode, I.ItemName, I.U_Grade AS [Grade],
    CAST(I.U_Thick AS DECIMAL(18,2)) AS [Thickness],
    CAST(I.U_Width AS INT) AS [Width],
    CAST(I.U_Length1 AS INT) AS [Length],
    T0.BatchNum, T0.U_CoilNo,
    T0.Quantity, T0.WhsCode,
    B.BinCode AS [Bin],
    T0.BaseNum AS [GRPONo],
    (SELECT MAX(A.DocDate) FROM OPDN A WHERE A.DocNum = T0.BaseNum) AS [GRPODate]
FROM OIBT T0
INNER JOIN OITM I ON T0.ItemCode = I.ItemCode
LEFT JOIN OBTN BT ON BT.ItemCode = T0.ItemCode AND BT.DistNumber = T0.BatchNum
LEFT JOIN OBBQ Q  ON Q.ItemCode = BT.ItemCode AND Q.SnBMDAbs = BT.AbsEntry AND Q.WhsCode = T0.WhsCode
LEFT JOIN OBIN B  ON B.AbsEntry = Q.BinAbs
WHERE T0.Quantity > 0 AND T0.WhsCode IN ('38','22')
ORDER BY T0.ItemCode, T0.BatchNum, B.BinCode
"@

$plateSql = @"
WITH InventoryData AS (
    SELECT T.ItemCode, I.ItemName, I.U_Grade, I.U_Thick, I.U_Width, I.U_Length1,
        (T.InQty - T.OutQty) AS NetQty,
        CASE
            WHEN T.TransType = 59 THEN ISNULL(GR.U_OutPcs, ISNULL(IGN1.U_Nos, 0))
            WHEN T.TransType = 60 THEN -1 * ISNULL(GI.U_OutPcs, ISNULL(IGE1.U_Nos, 0))
            WHEN T.TransType = 14 THEN -1 * ISNULL(RIN1.U_Nos, 0)
            WHEN T.OutQty > 0 THEN -1 * ISNULL(COALESCE(INV1.U_Nos, PCH1.U_Nos, DLN1.U_Nos, RDN1.U_Nos, RIN1.U_Nos, RPC1.U_Nos, WTR1.U_Nos), 0)
            ELSE ISNULL(COALESCE(INV1.U_Nos, PCH1.U_Nos, DLN1.U_Nos, RDN1.U_Nos, RIN1.U_Nos, RPC1.U_Nos, WTR1.U_Nos), 0)
        END AS RowNos
    FROM OINM T
    INNER JOIN OITM I ON I.ItemCode = T.ItemCode
    LEFT JOIN (SELECT U_GRDocEty, U_OPItem, U_GRLineNo, SUM(CAST(U_ActRcpt AS DECIMAL(19,6))) AS U_OutPcs FROM [@SMS_SHE1] GROUP BY U_GRDocEty, U_OPItem, U_GRLineNo) GR
        ON T.TransType = 59 AND GR.U_GRDocEty = T.CreatedBy AND GR.U_OPItem = T.ItemCode AND GR.U_GRLineNo = T.DocLineNum
    LEFT JOIN (SELECT U_GIDocEty, U_OPItem, U_GILineNo, SUM(CAST(U_InputPcs AS DECIMAL(19,6))) AS U_OutPcs FROM [@SMS_SHE1] GROUP BY U_GIDocEty, U_OPItem, U_GILineNo) GI
        ON T.TransType = 60 AND GI.U_GIDocEty = T.CreatedBy AND GI.U_OPItem = T.ItemCode AND GI.U_GILineNo = T.DocLineNum
    LEFT JOIN INV1 ON T.TransType = 13 AND INV1.DocEntry = T.CreatedBy AND INV1.LineNum = T.DocLineNum
    LEFT JOIN PCH1 ON T.TransType = 18 AND PCH1.DocEntry = T.CreatedBy AND PCH1.LineNum = T.DocLineNum
    LEFT JOIN DLN1 ON T.TransType = 15 AND DLN1.DocEntry = T.CreatedBy AND DLN1.LineNum = T.DocLineNum
    LEFT JOIN RDN1 ON T.TransType = 16 AND RDN1.DocEntry = T.CreatedBy AND RDN1.LineNum = T.DocLineNum
    LEFT JOIN RIN1 ON T.TransType = 14 AND RIN1.DocEntry = T.CreatedBy AND RIN1.LineNum = T.DocLineNum
    LEFT JOIN RPC1 ON T.TransType = 19 AND RPC1.DocEntry = T.CreatedBy AND RPC1.LineNum = T.DocLineNum
    LEFT JOIN WTR1 ON T.TransType = 67 AND WTR1.DocEntry = T.CreatedBy AND WTR1.LineNum = T.DocLineNum
    LEFT JOIN IGN1 ON T.TransType = 59 AND IGN1.DocEntry = T.CreatedBy AND IGN1.LineNum = T.DocLineNum
    LEFT JOIN IGE1 ON T.TransType = 60 AND IGE1.DocEntry = T.CreatedBy AND IGE1.LineNum = T.DocLineNum
)
SELECT ItemCode, MAX(ItemName) AS ItemName, MAX(U_Grade) AS Grade,
    MAX(U_Thick) AS Thickness, MAX(U_Width) AS Width, MAX(U_Length1) AS Length,
    SUM(NetQty) AS TotalQty, SUM(RowNos) AS TotalNos
FROM InventoryData
WHERE ItemCode LIKE 'HRC%'
GROUP BY ItemCode
HAVING SUM(NetQty) <> 0 OR SUM(RowNos) <> 0
ORDER BY ItemCode
"@


# ---- MIS queries: mirror the SAP saved queries behind the Excel pivot ----
#   "Warehouse Stock Report" / "Sales Report" / "Purchase" (Query Manager, category 40)
#   Category CASEs are copied verbatim so the dashboard reconciles to the pivot.

$misStockSql = @"
SELECT T0.ItemCode, T1.ItemName, T2.ItmsGrpNam AS ItemGroup,
  CASE
    WHEN T2.ItmsGrpNam IN ('SAIL H.R.PLATE','SAIL HR COIL','SAIL HR SHEET','HR Coil') THEN '1. SAIL HR'
    WHEN T2.ItmsGrpNam = 'SAIL STRUCTURAL' THEN '2. SAIL STRUCTURALS'
    WHEN T2.ItmsGrpNam = 'RINL TMT' THEN '3. RINL TMT'
    WHEN T2.ItmsGrpNam = 'RINL STRL' THEN '4. RINL STRL'
    WHEN T2.ItmsGrpNam IN ('JSW H.R.COIL','JSW H.R.PLATE','JSW H.R.SHEET') THEN '5. JSW HR'
    WHEN T2.ItmsGrpNam IN ('TATA H.R.COIL','TATA H.R.SHEET') THEN '6. TATA HR'
    WHEN T2.ItmsGrpNam IN ('SAIL TMT','SAIL TMT - SEQR') THEN '7. SAIL TMT'
    WHEN T2.ItmsGrpNam IN ('Consumables','Consumables-Critical') THEN '9. CONSUMABLES'
    ELSE '8. OTHERS'
  END AS Category,
  T0.WhsCode, T3.WhsName,
  T0.OnHand, T0.IsCommited AS Committed, T0.OnOrder,
  T1.U_Grade AS Grade, T1.U_Thick AS Thickness, T1.U_Width AS Width, T1.U_Length1 AS Length
FROM OITW T0
INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
INNER JOIN OITB T2 ON T1.ItmsGrpCod = T2.ItmsGrpCod
LEFT  JOIN OWHS T3 ON T3.WhsCode = T0.WhsCode
WHERE T0.ItemCode NOT IN ('a-item') AND T0.OnHand > 0
ORDER BY T0.WhsCode, T0.OnHand DESC
"@

# Sales: A/R invoices, same filters as the "Sales Report" query. Rolling from the
# start of the previous financial year so the page can show FY-on-FY.
$misSalesSql = @"
SELECT Category, ItemGroup, ItemCode, ItemName, CardName, Yr, Mth, JW, ProcessType,
       SUM(Qty) AS Qty, SUM(Value) AS Value, COUNT(DISTINCT DocEntry) AS Docs
FROM (
  SELECT
    CASE WHEN ISNULL(T0.U_VSPJWTX,'')='Job Work' THEN 'Y' ELSE 'N' END AS JW,
    ISNULL(T0.U_PROCESSTYPE,'') AS ProcessType,
    CASE
      WHEN T9.ItmsGrpNam IN ('SAIL H.R.PLATE','SAIL HR SHEET','SAIL HR COIL','HR Coil') THEN '1. SAIL HR'
      WHEN T9.ItmsGrpNam = 'SAIL STRUCTURAL' THEN '2. SAIL STRUCTURAL'
      WHEN T9.ItmsGrpNam = 'RINL TMT'  THEN '3. RINL TMT'
      WHEN T9.ItmsGrpNam = 'RINL STRL' THEN '4. RINL STRL'
      WHEN T9.ItmsGrpNam IN ('JSW H.R.SHEET','JSW H.R.COIL','SAIL TMT - SEQR','SAIL TMT','SAIL P.M.PLATE','OTHERS','TATA H.R.SHEET') THEN '5. OTHERS'
      ELSE NULL
    END AS Category,
    T9.ItmsGrpNam AS ItemGroup, T1.ItemCode, T6.ItemName, T0.CardName,
    YEAR(T0.DocDate) AS Yr, MONTH(T0.DocDate) AS Mth,
    T1.Quantity AS Qty, T1.LineTotal AS Value, T0.DocEntry
  FROM OINV T0
  INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry
  LEFT  JOIN OITM T6 ON T1.ItemCode = T6.ItemCode
  LEFT  JOIN OITB T9 ON T6.ItmsGrpCod = T9.ItmsGrpCod
  WHERE T1.BaseType <> '13' AND T0.Canceled = 'N' AND T0.GSTTranTyp <> 'GD'
    AND T1.Quantity > 0 AND T0.DocDate >= @StartDate
) X
WHERE Category IS NOT NULL
GROUP BY Category, ItemGroup, ItemCode, ItemName, CardName, Yr, Mth, JW, ProcessType
ORDER BY Yr, Mth, Category
"@

# Purchase: A/P invoices, same filters + line-level U_MOU categorisation as the "Purchase" query.
$misPurchSql = @"
SELECT Category, ItemGroup, ItemCode, ItemName, CardName, Yr, Mth,
       SUM(Qty) AS Qty, SUM(Value) AS Value, COUNT(DISTINCT DocEntry) AS Docs
FROM (
  SELECT
    CASE
      WHEN T1.U_MOU IN ('SAIL H.R.PLATE','SAIL HR COIL') THEN '1. SAIL HR'
      WHEN T1.U_MOU = 'SAIL STRUCTURAL' THEN '2. SAIL STRUCTURALS'
      WHEN T1.U_MOU = 'RINL TMT'  THEN '3. RINL TMT'
      WHEN T1.U_MOU = 'RINL STRL' THEN '4. RINL STRL'
      WHEN T1.U_MOU IN ('OTHERS','RINL ROUNDS','SAIL P.M.PLATE','TATA','TENDER','JSW H.R.COIL','SAIL TMT') THEN '5. OTHERS'
      ELSE LTRIM(RTRIM(T1.U_MOU))
    END AS Category,
    LTRIM(RTRIM(T1.U_MOU)) AS ItemGroup, T1.ItemCode, T6.ItemName, T0.CardName,
    YEAR(T0.DocDate) AS Yr, MONTH(T0.DocDate) AS Mth,
    T1.Quantity AS Qty, T1.LineTotal AS Value, T0.DocEntry
  FROM OPCH T0
  INNER JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry
  LEFT  JOIN OITM T6 ON T1.ItemCode = T6.ItemCode
  WHERE T1.BaseType <> '18' AND T0.Canceled = 'N' AND T0.GSTTranTyp <> 'GD'
    AND T1.Quantity > 0 AND LTRIM(RTRIM(ISNULL(T1.U_MOU,''))) <> ''
    AND T0.DocDate >= @StartDate
) X
GROUP BY Category, ItemGroup, ItemCode, ItemName, CardName, Yr, Mth
ORDER BY Yr, Mth, Category
"@

# ---- Job work (customer-owned material processed at G2 and returned) ----
# Flow through warehouse 45 by transaction type, party and month.
$misJwFlowSql = @"
SELECT YEAR(M.DocDate) AS Yr, MONTH(M.DocDate) AS Mth,
  CASE M.TransType WHEN 20 THEN 'Received (GRPO)' WHEN 67 THEN 'Received (Transfer)' WHEN 59 THEN 'Processed (GR)'
                   WHEN 60 THEN 'Issued to process (GI)' WHEN 15 THEN 'Dispatched (Delivery)' WHEN 14 THEN 'Returned (Credit)'
                   ELSE CAST(M.TransType AS NVARCHAR(10)) END AS Kind,
  ISNULL(M.CardName,'') AS CardName, SUM(M.InQty) AS InQty, SUM(M.OutQty) AS OutQty, COUNT(*) AS Lines
FROM OINM M
WHERE M.Warehouse='45' AND M.DocDate >= @StartDate
GROUP BY YEAR(M.DocDate), MONTH(M.DocDate), M.TransType, M.CardName
ORDER BY Yr, Mth, Kind
"@

# Job-work deliveries and invoices (tagged by the VSP add-on), plus delivered-not-invoiced.
$misJwDocsSql = @"
SELECT 'Delivery' AS Kind, H.CardName, ISNULL(H.U_PROCESSTYPE,'') AS ProcessType, YEAR(H.DocDate) Yr, MONTH(H.DocDate) Mth,
       SUM(L.Quantity) AS Qty, SUM(L.LineTotal) AS Value, COUNT(DISTINCT H.DocEntry) AS Docs,
       SUM(CASE WHEN L.TrgetEntry IS NULL AND L.LineStatus='O' THEN L.Quantity ELSE 0 END) AS OpenQty
FROM ODLN H JOIN DLN1 L ON L.DocEntry=H.DocEntry
WHERE H.DocDate >= @StartDate AND H.CANCELED='N' AND ISNULL(H.U_VSPJWTX,'')='Job Work'
GROUP BY H.CardName, H.U_PROCESSTYPE, YEAR(H.DocDate), MONTH(H.DocDate)
UNION ALL
SELECT 'Invoice', H.CardName, ISNULL(H.U_PROCESSTYPE,''), YEAR(H.DocDate), MONTH(H.DocDate),
       SUM(L.Quantity), SUM(L.LineTotal), COUNT(DISTINCT H.DocEntry), 0
FROM OINV H JOIN INV1 L ON L.DocEntry=H.DocEntry
WHERE H.DocDate >= @StartDate AND H.CANCELED='N' AND H.GSTTranTyp<>'GD' AND ISNULL(H.U_VSPJWTX,'')='Job Work'
GROUP BY H.CardName, H.U_PROCESSTYPE, YEAR(H.DocDate), MONTH(H.DocDate)
ORDER BY Yr, Mth, Kind
"@

# What is lying in the job-work warehouse right now, by batch and owner.
$misJwBalanceSql = @"
SELECT B.ItemCode, I.ItemName, B.BatchNum, ISNULL(B.U_VendCode,'') AS OwnerCode, ISNULL(C.CardName,'') AS Owner,
       B.Quantity, B.InDate, DATEDIFF(day, B.InDate, GETDATE()) AS AgeDays
FROM OIBT B JOIN OITM I ON I.ItemCode=B.ItemCode
LEFT JOIN OCRD C ON C.CardCode=B.U_VendCode
WHERE B.WhsCode='45' AND B.Quantity<>0
ORDER BY B.Quantity DESC
"@

# ---- Wheels India job-work stock (WH45, this customer only) ----
$wheelsSql = @"
;WITH WIRItems AS (
    SELECT W.ItemCode, I.ItemName, I.ItmsGrpCod, I.U_Grade, I.U_Thick, I.U_Width, I.U_Length1, W.OnHand
    FROM OITW W INNER JOIN OITM I ON I.ItemCode = W.ItemCode
    WHERE W.WhsCode = '45' AND W.OnHand > 0
      AND (I.ItemName LIKE '%WIR0%' OR I.ItemName LIKE '%WIR %' OR I.ItemName LIKE '%WIR')
),
Batch AS (  -- one representative batch per item (the largest), for coil/mother/age reference
    SELECT B.ItemCode, B.BatchNum, B.U_CoilNo, B.U_Mothercoil, B.InDate,
           ROW_NUMBER() OVER (PARTITION BY B.ItemCode ORDER BY B.Quantity DESC) AS rn,
           COUNT(*) OVER (PARTITION BY B.ItemCode) AS BatchCount
    FROM OIBT B WHERE B.WhsCode='45' AND B.Quantity <> 0
)
SELECT
    WI.ItemCode, WI.ItemName, G.ItmsGrpNam AS ItemGroup,
    '' AS OwnerCode, '' AS OwnerName,
    ISNULL(BX.BatchNum,'') AS BatchNum, ISNULL(BX.U_CoilNo,'') AS CoilNo, ISNULL(BX.U_Mothercoil,'') AS Mothercoil,
    ISNULL(WI.U_Grade,'') AS Grade, WI.U_Thick AS Thick, WI.U_Width AS Width, WI.U_Length1 AS Length,
    WI.OnHand AS Quantity,          -- authoritative on-hand (reconciles to SAP)
    ISNULL(BX.BatchCount,0) AS BatchCount,
    DATEDIFF(day, BX.InDate, GETDATE()) AS AgeDays
FROM WIRItems WI
LEFT JOIN Batch BX ON BX.ItemCode = WI.ItemCode AND BX.rn = 1
LEFT JOIN OITB G ON G.ItmsGrpCod = WI.ItmsGrpCod
ORDER BY WI.ItemName
"@

# ---- Wheels India coil traceability (every WH45 WIR batch, with lifecycle status) ----
# One row per batch ever in WH45: current on-hand (In Whs), plus how much of it was
# Delivered (IBT1 BaseType 15 = Delivery note, OR 13 = A/R Invoice billed directly
# with no delivery note - both post the goods movement) or Processed/cut (goods issue,
# IBT1 BaseType 60). Both are NET of direction: an OUT movement (Direction=1) adds,
# an IN movement (Direction=2) subtracts, so a reversed/returned delivery or a
# cancelled goods-issue nets back to zero instead of being double-counted.
# Source coil = U_CoilNo (SAP keeps it identical on the coil and
# every derivative). The portal classifies Coil/Plate/Plate-Scrap/Strip/Strip-Scrap by
# item name and picks the status from these three quantities.
$wtSql = @"
SELECT
  B.ItemCode,
  I.ItemName,
  B.BatchNum,
  ISNULL(B.U_CoilNo,'')     AS CoilNo,
  ISNULL(B.U_Mothercoil,'') AS Mothercoil,
  ISNULL(I.U_Grade,'')      AS Grade,
  I.U_Thick                 AS Thick,
  I.U_Width                 AS Width,
  I.U_Length1               AS Length,
  CAST(B.Quantity AS DECIMAL(19,3))          AS InWhsQty,
  CAST(ISNULL(M.DelQty,0) AS DECIMAL(19,3))  AS DeliveredQty,
  CAST(ISNULL(M.IssQty,0) AS DECIMAL(19,3))  AS ProcessedQty
FROM OIBT B
INNER JOIN OITM I ON I.ItemCode = B.ItemCode
LEFT JOIN (
  SELECT ItemCode, BatchNum,
    SUM(CASE WHEN BaseType IN (13,15) THEN (CASE WHEN Direction=1 THEN Quantity ELSE -Quantity END) ELSE 0 END) AS DelQty,
    SUM(CASE WHEN BaseType=60 THEN (CASE WHEN Direction=1 THEN Quantity ELSE -Quantity END) ELSE 0 END) AS IssQty
  FROM IBT1 WHERE WhsCode='45' GROUP BY ItemCode, BatchNum
) M ON M.ItemCode = B.ItemCode AND M.BatchNum = B.BatchNum
WHERE B.WhsCode = '45'
  AND (I.ItemName LIKE '%WIR0%' OR I.ItemName LIKE '%WIR %' OR I.ItemName LIKE '%WIR')
  AND (B.Quantity > 0 OR M.DelQty <> 0 OR M.IssQty <> 0)
ORDER BY B.U_CoilNo, B.BatchNum
"@

# ---- Production plan schedule (recursive SMS schedule flow: coil -> cut plates, all stages) ----
# Feeds the Wheels India "Production Schedule" portal tab. Returns ALL production; the tab
# defaults the view to Open + WIR grades. One row per output line per stage.
# [Next Schedule No] = 0 or NULL means the plate is a final deliverable (not fed to a further stage).
$wheelsSchedSql = @'
;WITH ProductionFlow AS
(
    SELECT
        H.DocNum AS [Schedule Master No],
        H.DocEntry AS [Schedule Master Entry],
        CAST(H.U_DocDate   AS DATE) AS [Schedule Master Date],       -- schedule document date
        CAST(H.U_WIO_Date  AS DATE) AS [Schedule Master WIO Date],   -- Wheels India ORDER date (plan received)
        CAST(H.U_WIPPDate  AS DATE) AS [Schedule Master WI Production Plan Date], -- WI PRODUCTION PLANNED date
        CAST(OH.CreateDate AS DATE) AS [Program Date],               -- program creation date (fallback)
        1 AS [Stage No],
        CAST(S.U_IPItem AS NVARCHAR(100)) AS [Input Item Code],
        CAST(S.U_IPDesrp AS NVARCHAR(254)) AS [Input Item Name],
        TRY_CONVERT(DECIMAL(19,6), S.U_IOnStock) AS [Input Qty],
        TRY_CONVERT(DECIMAL(19,6), S.U_ProdQty) AS [Production Qty],
        OH.DocNum AS [Output Schedule No],
        CAST(OH.U_DocDate  AS DATE) AS [Schedule Output Date],
        CAST(OH.U_WIO_Date AS DATE) AS [Schedule Output WIO Date],
        CAST(OH.U_WIPPDate AS DATE) AS [Schedule Output WI Production Plan Date],
        CAST(O.U_OPItem AS NVARCHAR(100)) AS [Output Item Code],
        CAST(O.U_OPDesrp AS NVARCHAR(254)) AS [Output Item Name],
        TRY_CONVERT(DECIMAL(19,6), O.U_OPQty) AS [Output Qty],
        CAST(O.U_OPThick AS NVARCHAR(50)) AS [Thickness],
        CAST(O.U_OPWidth AS NVARCHAR(50)) AS [Output Width],
        CAST(O.U_OPLength AS NVARCHAR(50)) AS [Output Length],
        TRY_CONVERT(INT, O.U_OutPcs) AS [Actual Pcs],           -- No. of pieces (output)
        CAST(CAST(O.U_Remarks AS NVARCHAR(500)) AS NVARCHAR(500)) AS [Remarks],
        TRY_CONVERT(INT, O.U_SHONum) AS [Next Schedule No],
        CAST(CASE WHEN OH.Status='O' THEN 'Open' WHEN OH.Status='C' THEN 'Close' ELSE OH.Status END AS NVARCHAR(20)) AS [Status],
        CAST('|' + CAST(OH.DocNum AS NVARCHAR(20)) + '|' AS NVARCHAR(MAX)) AS [Document Path]
    FROM [@SMS_OSCH] H
    INNER JOIN [@SMS_SCH1] S ON H.DocEntry = S.DocEntry
    INNER JOIN [@SMS_OSHO] OH ON TRY_CONVERT(INT, S.U_SHONum) = OH.DocNum
    INNER JOIN [@SMS_SHO1] O ON OH.DocEntry = O.DocEntry
    UNION ALL
    SELECT
        P.[Schedule Master No],
        P.[Schedule Master Entry],
        P.[Schedule Master Date],
        P.[Schedule Master WIO Date],
        P.[Schedule Master WI Production Plan Date],
        CAST(OH2.CreateDate AS DATE) AS [Program Date],
        P.[Stage No] + 1 AS [Stage No],
        P.[Output Item Code] AS [Input Item Code],
        P.[Output Item Name] AS [Input Item Name],
        P.[Output Qty] AS [Input Qty],
        CAST(NULL AS DECIMAL(19,6)) AS [Production Qty],
        OH2.DocNum AS [Output Schedule No],
        CAST(OH2.U_DocDate  AS DATE) AS [Schedule Output Date],
        CAST(OH2.U_WIO_Date AS DATE) AS [Schedule Output WIO Date],
        CAST(OH2.U_WIPPDate AS DATE) AS [Schedule Output WI Production Plan Date],
        CAST(O2.U_OPItem AS NVARCHAR(100)) AS [Output Item Code],
        CAST(O2.U_OPDesrp AS NVARCHAR(254)) AS [Output Item Name],
        TRY_CONVERT(DECIMAL(19,6), O2.U_OPQty) AS [Output Qty],
        CAST(O2.U_OPThick AS NVARCHAR(50)) AS [Thickness],
        CAST(O2.U_OPWidth AS NVARCHAR(50)) AS [Output Width],
        CAST(O2.U_OPLength AS NVARCHAR(50)) AS [Output Length],
        TRY_CONVERT(INT, O2.U_OutPcs) AS [Actual Pcs],
        CAST(CAST(O2.U_Remarks AS NVARCHAR(500)) AS NVARCHAR(500)) AS [Remarks],
        TRY_CONVERT(INT, O2.U_SHONum) AS [Next Schedule No],
        CAST(CASE WHEN OH2.Status='O' THEN 'Open' WHEN OH2.Status='C' THEN 'Close' ELSE OH2.Status END AS NVARCHAR(20)) AS [Status],
        CAST(P.[Document Path] + CAST(OH2.DocNum AS NVARCHAR(20)) + '|' AS NVARCHAR(MAX)) AS [Document Path]
    FROM ProductionFlow P
    INNER JOIN [@SMS_OSHO] OH2 ON P.[Next Schedule No] = OH2.DocNum
    INNER JOIN [@SMS_SHO1] O2 ON OH2.DocEntry = O2.DocEntry
    WHERE P.[Next Schedule No] IS NOT NULL
      AND CHARINDEX('|' + CAST(OH2.DocNum AS NVARCHAR(20)) + '|', P.[Document Path]) = 0
)
SELECT
    [Schedule Master No],[Schedule Master Entry],
    [Schedule Master Date],[Schedule Master WIO Date],[Schedule Master WI Production Plan Date],[Program Date],
    [Stage No],
    [Input Item Code],[Input Item Name],[Input Qty],[Production Qty],
    [Output Schedule No],[Schedule Output Date],[Schedule Output WIO Date],[Schedule Output WI Production Plan Date],
    [Output Item Code],[Output Item Name],[Output Qty],
    [Thickness],[Output Width],[Output Length],[Actual Pcs],[Remarks],[Next Schedule No],[Status]
FROM ProductionFlow
ORDER BY [Schedule Master No],[Stage No],[Output Schedule No],[Output Item Code]
OPTION (MAXRECURSION 100);
'@

function Invoke-SqlDated($sql, [string]$startDate){
  $conn=New-Object System.Data.SqlClient.SqlConnection $connStr
  $conn.Open()
  $cmd=$conn.CreateCommand(); $cmd.CommandText=$sql; $cmd.CommandTimeout=180
  $null=$cmd.Parameters.AddWithValue("@StartDate",$startDate)
  $da=New-Object System.Data.SqlClient.SqlDataAdapter $cmd
  $dt=New-Object System.Data.DataTable; [void]$da.Fill($dt)
  $conn.Close(); return ,$dt
}

# DataTable -> CSV text (UTF-8, no BOM, quoted only when needed, dates as yyyy-MM-dd)
function ConvertTo-CsvText([System.Data.DataTable]$dt){
  $cols=@($dt.Columns | ForEach-Object { $_.ColumnName })
  $esc={ param($s) if($null -eq $s){return ''}; if($s -match '[",\r\n]'){ '"'+($s -replace '"','""')+'"' } else { $s } }
  $sb=New-Object System.Text.StringBuilder
  [void]$sb.AppendLine((($cols | ForEach-Object { & $esc $_ }) -join ','))
  foreach($r in $dt.Rows){
    $vals=foreach($c in $cols){ $v=$r[$c]
      if($v -is [DBNull] -or $null -eq $v){''} elseif($v -is [DateTime]){$v.ToString('yyyy-MM-dd')}
      elseif($v -is [decimal] -or $v -is [double]){ ([decimal]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }
      else{ & $esc ([string]$v) } }
    [void]$sb.AppendLine(($vals -join ','))
  }
  return $sb.ToString()
}

# ---- Weighbridge / dispatch feed (reconstructed from @VSPLWB) ----
# Reproduces the old dispatch/data.csv exactly: same 19 columns, same order.
#   Net Weight = U_VSPWNT/1000 (12 dp)  |  SO Qty = SUM(RDR1.Quantity) for the SO
#   Ship To Name = ORDR.ShipToCode      |  Invoice No/Time/Branch via the SO's A/R invoice
#   Invoice Time is zero-padded to HH:MM and OutTime->Invoice gap parsed off the real H:M
#   (so single-digit minutes like "16:5" -> "16:05" / "41 Min", matching the live feed).
# Rolling 1-year window, newest first.
$wbSql = @"
SELECT
  W.CreateDate                         AS [CreateDate],
  W.U_VSPCN                            AS [U_VSPCN],
  O2.ShipToCode                        AS [Ship To Name],
  NULLIF(W.U_VSPWBITM,0)               AS [Report Time],
  W.CreateTime                         AS [In Time],
  NULLIF(W.U_VSPWBOTM,0)               AS [Out Time],
  NULLIF(W.U_VSPLDITM,0)               AS [Time Taken],
  CAST(ISNULL(W.U_VSPWNT,0)/1000.0 AS DECIMAL(19,12)) AS [Net Weight],
  CAST(SOQ.SOQty AS DECIMAL(19,6))     AS [SO Qty],
  W.U_VSPV                             AS [U_VSPV],
  W.U_VSPLSLEM                         AS [Sales Employee Name],
  INV.DocNum                           AS [Invoice No],
  CASE WHEN INV.U_Inv_time IS NOT NULL AND CHARINDEX(':', INV.U_Inv_time) > 0
       THEN LEFT(INV.U_Inv_time, CHARINDEX(':', INV.U_Inv_time)-1)
            + ':' + RIGHT('0' + LTRIM(RTRIM(SUBSTRING(INV.U_Inv_time, CHARINDEX(':', INV.U_Inv_time)+1, 10))), 2)
       ELSE INV.U_Inv_time END         AS [Invoice Time],
  CASE WHEN INV.U_Inv_time IS NOT NULL
        AND CHARINDEX(':', INV.U_Inv_time) > 0
        AND ISNULL(W.U_VSPWBOTM,0) > 0
        AND TRY_CONVERT(INT, LTRIM(RTRIM(LEFT(INV.U_Inv_time, CHARINDEX(':',INV.U_Inv_time)-1)))) IS NOT NULL
        AND TRY_CONVERT(INT, LTRIM(RTRIM(SUBSTRING(INV.U_Inv_time, CHARINDEX(':',INV.U_Inv_time)+1, 10)))) IS NOT NULL
       THEN CAST(
              ( TRY_CONVERT(INT, LTRIM(RTRIM(LEFT(INV.U_Inv_time, CHARINDEX(':',INV.U_Inv_time)-1))))*60
                + TRY_CONVERT(INT, LTRIM(RTRIM(SUBSTRING(INV.U_Inv_time, CHARINDEX(':',INV.U_Inv_time)+1, 10)))) )
              - ((W.U_VSPWBOTM/100)*60 + (W.U_VSPWBOTM%100))
            AS NVARCHAR(10)) + ' Min'
       ELSE '' END                     AS [OutTime_To_InvoiceTime],
  W.U_VSPLSDE                          AS [SO DocEntry No],
  W.U_VSPLORNO                         AS [SO No],
  CASE WHEN ISNULL(NULLIF(LTRIM(RTRIM(INV.U_Branch)),''),'CHN')='CHN' THEN 'CHENNAI' ELSE INV.U_Branch END AS [Branch],
  W.U_VSPSTATUS                        AS [U_VSPSTATUS],
  W.U_VSPLWBST                         AS [U_VSPLWBST]
FROM [@VSPLWB] W
LEFT JOIN ORDR O2 ON O2.DocEntry = W.U_VSPLSDE
OUTER APPLY (SELECT SUM(L.Quantity) AS SOQty FROM RDR1 L WHERE L.DocEntry = W.U_VSPLSDE) SOQ
OUTER APPLY (SELECT TOP 1 H.DocNum, H.U_Inv_time, H.U_Branch
             FROM OINV H INNER JOIN INV1 L ON L.DocEntry = H.DocEntry
             WHERE L.BaseType = 17 AND L.BaseEntry = W.U_VSPLSDE
             ORDER BY H.DocEntry DESC) INV
WHERE ISNULL(W.Canceled,'N') = 'N'
  AND W.CreateDate >= DATEADD(YEAR, -1, CAST(GETDATE() AS DATE))
ORDER BY W.CreateDate DESC, W.DocEntry DESC
"@

# ======================= RUN =======================
# ---- lock: refuse to overlap with a still-running instance (stale after 10 min) ----
$LockPath="$env:TEMP\refresh_stock.lock"
if(Test-Path $LockPath){
  $age=(Get-Date)-(Get-Item $LockPath).LastWriteTime
  if($age.TotalMinutes -lt 10){ Log "Another instance is running (lock age $([int]$age.TotalSeconds)s). Exiting."; exit 0 }
  Remove-Item $LockPath -Force
}
New-Item $LockPath -ItemType File -Force | Out-Null

# make sure the stock\ subfolder exists
if(-not (Test-Path $StockDir)){ New-Item $StockDir -ItemType Directory -Force | Out-Null }

# Financial year starts 1 April. Pull from the start of the PREVIOUS FY.
$now=Get-Date; $fy= if($now.Month -ge 4){$now.Year}else{$now.Year-1}
$StartDate=(Get-Date -Year ($fy-1) -Month 4 -Day 1).ToString('yyyy-MM-dd')

try{
  Log "Querying coils..."
  $coils = Invoke-Sql $coilSql
  Log "  $($coils.Rows.Count) coil rows"

  Log "Querying plates..."
  $plates = Invoke-Sql $plateSql
  Log "  $($plates.Rows.Count) plate rows"

  # ---- build coils.csv (add Brand, tidy GRPODate) ----
  $coilOut = foreach($r in $coils.Rows){
    [PSCustomObject]@{
      ItemCode=$r.ItemCode; ItemName=$r.ItemName; Brand=(Get-Brand $r.ItemName)
      Grade=$r.Grade; Thickness=$r.Thickness; Width=$r.Width; Length=$r.Length
      BatchNum=$r.BatchNum; CoilNo=$r.U_CoilNo; Quantity=$r.Quantity
      WhsCode=$r.WhsCode; Bin=$r.Bin; GRPONo=$r.GRPONo
      GRPODate=if($r.GRPODate -is [DateTime]){ $r.GRPODate.ToString('yyyy-MM-dd') } else { '' }
    }
  }
  # ---- build plates.csv ----
  $plateOut = foreach($r in $plates.Rows){
    [PSCustomObject]@{
      ItemCode=$r.ItemCode; ItemName=$r.ItemName; Brand=(Get-Brand $r.ItemName)
      Grade=$r.Grade; Thickness=$r.Thickness; Width=$r.Width; Length=$r.Length
      TotalQty=$r.TotalQty; TotalNos=$r.TotalNos
    }
  }

  # write WITHOUT BOM (dashboard strips BOM anyway, but keep it clean)
  $enc=New-Object System.Text.UTF8Encoding($false)
  [IO.File]::WriteAllText("$StockDir\coils.csv",  ($coilOut  | ConvertTo-Csv -NoTypeInformation) -join "`r`n", $enc)
  [IO.File]::WriteAllText("$StockDir\plates.csv", ($plateOut | ConvertTo-Csv -NoTypeInformation) -join "`r`n", $enc)
  Log "CSVs written."

  # ---- MIS: stock / sales / purchase ----
  Log "Querying MIS stock...";    $misStock=Invoke-Sql $misStockSql;                 Log "  $($misStock.Rows.Count) rows"
  Log "Querying MIS sales...";    $misSales=Invoke-SqlDated $misSalesSql $StartDate;  Log "  $($misSales.Rows.Count) rows"
  Log "Querying MIS purchase..."; $misPurch=Invoke-SqlDated $misPurchSql $StartDate;  Log "  $($misPurch.Rows.Count) rows"
  Log "Querying job work...";     $jwFlow=Invoke-SqlDated $misJwFlowSql $StartDate;    $jwDocs=Invoke-SqlDated $misJwDocsSql $StartDate;  $jwBal=Invoke-Sql $misJwBalanceSql
  Log "  flow $($jwFlow.Rows.Count) - docs $($jwDocs.Rows.Count) - balance $($jwBal.Rows.Count) rows"
  Log "Querying Wheels India (WH45)..."; $wheels=Invoke-Sql $wheelsSql; Log "  $($wheels.Rows.Count) rows"
  Log "Querying Wheels coil traceability..."; $wtrace=$null
  try{ $wtrace=Invoke-Sql $wtSql; Log "  $($wtrace.Rows.Count) batch rows" }
  catch{ Log "WARN: traceability query failed ($($_.Exception.Message)) - keeping previous wheels_traceability.csv" }

  # Wheels production - daily machine-wise feed for the current calendar year.
  # The portal aggregates this into the 3-section report (Total / Wheels India / Monthly)
  # and applies the holiday-aware "last working day" logic, so no period maths is done here.
  #   Machines (MDS): MACHINE02 = CTL 01-06 MM, MACHINE01 = CTL 6-25 MM, MACHINE15 = Shearing 06-25 MM.
  #   [Total Qty] = all grades on these machines; [Wheels Qty] = WIR006/WIR025 only.
  #   NOTE: HRC-coded output items ARE counted - for the CTL lines the cut HR-coil output keeps its
  #   HRC item code and is genuine production (excluding it understated CTL 01-06). Only true scrap/
  #   offcut items (OTH000000000185 / OTH000000000218) and blank items are excluded.
  $wheelsDailySql = @'
SELECT
    CONVERT(varchar(10), COALESCE(TRY_CONVERT(DATE, NULLIF(T0.U_ExDate,'')), TRY_CONVERT(DATE, T0.U_SDate)), 23) AS [Date],
    T1.U_MachCode AS [Machine],
    MAX(T1.U_MachName) AS [Machine Name],
    CAST(SUM(ISNULL(TRY_CONVERT(DECIMAL(19,4), T1.U_OPQty),0)) AS DECIMAL(19,4)) AS [Total Qty],
    CAST(SUM(CASE WHEN T1.U_IPGrade IN ('WIR006','WIR025') THEN ISNULL(TRY_CONVERT(DECIMAL(19,4), T1.U_OPQty),0) ELSE 0 END) AS DECIMAL(19,4)) AS [Wheels Qty]
FROM [@SMS_OSHE] T0 INNER JOIN [@SMS_SHE1] T1 ON T0.DocEntry = T1.DocEntry
WHERE COALESCE(TRY_CONVERT(DATE, NULLIF(T0.U_ExDate,'')), TRY_CONVERT(DATE, T0.U_SDate)) >= DATEADD(YEAR, DATEDIFF(YEAR,0,GETDATE()),0)
  AND COALESCE(TRY_CONVERT(DATE, NULLIF(T0.U_ExDate,'')), TRY_CONVERT(DATE, T0.U_SDate)) <  DATEADD(YEAR, DATEDIFF(YEAR,0,GETDATE())+1,0)
  AND T1.U_MachCode IN ('MACHINE01 (MDS)','MACHINE02 (MDS)','MACHINE15 (MDS)')
  AND ISNULL(T1.U_OPItem,'') <> ''
  AND T1.U_OPItem NOT IN ('OTH000000000185','OTH000000000218')
GROUP BY COALESCE(TRY_CONVERT(DATE, NULLIF(T0.U_ExDate,'')), TRY_CONVERT(DATE, T0.U_SDate)), T1.U_MachCode
ORDER BY 1, 2
'@
  Log "Querying Wheels production (daily)..."; $wdaily=Invoke-Sql $wheelsDailySql; Log "  $($wdaily.Rows.Count) rows"

  Log "Querying production schedule..."; $wsched=Invoke-Sql $wheelsSchedSql; Log "  $($wsched.Rows.Count) rows"


  # production per machine (SMS shift-entry UDTs) - today + month-to-date, prime/defective/scrap
  $prodSql = @'
SELECT
    T1.U_Machcode AS [Machine Code],
    T1.U_MachName AS [Machine Name],
    SUM(CASE WHEN CAST(T0.U_SDate AS DATE) = CAST(GETDATE() AS DATE) AND T1.U_OPItem NOT IN ('OTH000000000218','OTH000000000185') THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_Actrcpt),0) ELSE 0 END) AS [Today Prime Nos],
    SUM(CASE WHEN CAST(T0.U_SDate AS DATE) = CAST(GETDATE() AS DATE) AND T1.U_OPItem = 'OTH000000000185' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_Actrcpt),0) ELSE 0 END) AS [Today Defective Nos],
    SUM(CASE WHEN CAST(T0.U_SDate AS DATE) = CAST(GETDATE() AS DATE) AND T1.U_OPItem = 'OTH000000000218' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_Actrcpt),0) ELSE 0 END) AS [Today Scrap Nos],
    SUM(CASE WHEN CAST(T0.U_SDate AS DATE) = CAST(GETDATE() AS DATE) AND T1.U_OPItem NOT IN ('OTH000000000218','OTH000000000185') THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) AS [Today Prime Qty],
    SUM(CASE WHEN CAST(T0.U_SDate AS DATE) = CAST(GETDATE() AS DATE) AND T1.U_OPItem = 'OTH000000000185' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) AS [Today Defective Qty],
    SUM(CASE WHEN CAST(T0.U_SDate AS DATE) = CAST(GETDATE() AS DATE) AND T1.U_OPItem = 'OTH000000000218' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) AS [Today Scrap Qty],
    SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem NOT IN ('OTH000000000218','OTH000000000185') THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_Actrcpt),0) ELSE 0 END) AS [Month Prime Nos],
    SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem = 'OTH000000000185' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_Actrcpt),0) ELSE 0 END) AS [Month Defective Nos],
    SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem = 'OTH000000000218' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_Actrcpt),0) ELSE 0 END) AS [Month Scrap Nos],
    SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem NOT IN ('OTH000000000218','OTH000000000185') THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) AS [Month Prime Qty],
    SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem = 'OTH000000000185' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) AS [Month Defective Qty],
    SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem = 'OTH000000000218' THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) AS [Month Scrap Qty]
FROM [@SMS_OSHE] T0
INNER JOIN [@SMS_SHE1] T1 ON T0.DocEntry = T1.DocEntry
WHERE NULLIF(LTRIM(RTRIM(T1.U_MachName)), '') IS NOT NULL
GROUP BY T1.U_Machcode, T1.U_MachName
HAVING SUM(CASE WHEN YEAR(T0.U_SDate) = YEAR(GETDATE()) AND MONTH(T0.U_SDate) = MONTH(GETDATE()) AND T1.U_OPItem NOT IN ('OTH000000000218','OTH000000000185') THEN ISNULL(TRY_CONVERT(DECIMAL(19,6),T1.U_OPQty),0) ELSE 0 END) <> 0
ORDER BY CASE T1.U_Machcode
    WHEN 'MACHINE01' THEN 1 WHEN 'MACHINE01 (MDS)' THEN 2 WHEN 'MACHINE01 (SLM)' THEN 3
    WHEN 'MACHINE02' THEN 4 WHEN 'MACHINE02 (MDS)' THEN 5 WHEN 'MACHINE 02 SLITTING' THEN 6
    WHEN 'MACHINE15 (MDS)' THEN 7 ELSE 99 END, T1.U_Machcode
'@
  Log "Querying production (SMS)..."
  $prodOk=$false
  try{ $prod=Invoke-Sql $prodSql; $prodOk=$true; Log "  $($prod.Rows.Count) rows" }
  catch{ Log "WARN: production (SMS) query failed ($($_.Exception.Message)) - keeping previous production.csv, refreshing everything else" }

  if(@($misStock.Rows).Count -eq 0){ throw "MIS stock query returned 0 rows - not overwriting dashboard data." }
  [IO.File]::WriteAllText("$StockDir\mis_stock.csv",    (ConvertTo-CsvText $misStock), $enc)
  [IO.File]::WriteAllText("$StockDir\mis_sales.csv",    (ConvertTo-CsvText $misSales), $enc)
  [IO.File]::WriteAllText("$StockDir\mis_purchase.csv", (ConvertTo-CsvText $misPurch), $enc)
  if($prodOk){ [IO.File]::WriteAllText("$StockDir\production.csv",   (ConvertTo-CsvText $prod), $enc) }
  [IO.File]::WriteAllText("$StockDir\mis_jobwork_flow.csv",    (ConvertTo-CsvText $jwFlow), $enc)
  [IO.File]::WriteAllText("$StockDir\mis_jobwork_docs.csv",    (ConvertTo-CsvText $jwDocs), $enc)
  [IO.File]::WriteAllText("$StockDir\mis_jobwork_balance.csv", (ConvertTo-CsvText $jwBal), $enc)
  # Wheels India data (same data repo, root)
  $WheelsDir = $RepoDir
  if(-not (Test-Path $WheelsDir)){ New-Item $WheelsDir -ItemType Directory -Force | Out-Null }
  [IO.File]::WriteAllText("$WheelsDir\wheels_stock.csv", (ConvertTo-CsvText $wheels), $enc)
  [IO.File]::WriteAllText("$WheelsDir\wheels_proddaily.csv",         (ConvertTo-CsvText $wdaily), $enc)
  [IO.File]::WriteAllText("$WheelsDir\wheels_prod_schedule.csv",     (ConvertTo-CsvText $wsched), $enc)
  if($wtrace -and $wtrace.Rows.Count -gt 0){ [IO.File]::WriteAllText("$WheelsDir\wheels_traceability.csv", (ConvertTo-CsvText $wtrace), $enc) }
  [IO.File]::WriteAllText("$WheelsDir\wheels_meta.json", (@{ generated=$now.ToString('yyyy-MM-ddTHH:mm:ss'); rows=$wheels.Rows.Count } | ConvertTo-Json -Compress), $enc)
  $meta=@{ generated=$now.ToString('yyyy-MM-ddTHH:mm:ss'); fyStart="$fy-04-01"; dataFrom=$StartDate;
           rows=@{ coils=$coils.Rows.Count; plates=$plates.Rows.Count; stock=$misStock.Rows.Count; sales=$misSales.Rows.Count; purchase=$misPurch.Rows.Count; jobwork=$jwDocs.Rows.Count } }
  [IO.File]::WriteAllText("$StockDir\mis_meta.json", ($meta | ConvertTo-Json -Compress), $enc)
  Log "MIS CSVs written."

  # ---- push (pull-first to avoid the race we hit on weighbridge) ----
  Push-Location $RepoDir
  # git writes warnings to stderr; under EAP=Stop + 2>&1 those become terminating errors.
  # Judge git by exit code only.
  $prevEap=$ErrorActionPreference; $ErrorActionPreference='Continue'
  try{
    git add coils.csv plates.csv mis_stock.csv mis_sales.csv mis_purchase.csv production.csv mis_jobwork_flow.csv mis_jobwork_docs.csv mis_jobwork_balance.csv mis_meta.json wheels_stock.csv wheels_proddaily.csv wheels_prod_schedule.csv wheels_traceability.csv wheels_meta.json 2>&1 | Out-Null
    $stamp=[DateTime]::Now.ToString('yyyy-MM-dd HH:mm')
    git commit -m "stock refresh $stamp" 2>&1 | Out-Null          # no-op if nothing changed
    # commit first, THEN rebase: -X theirs = keep our freshly generated CSVs on conflict,
    # --autostash tolerates any other unstaged edits sitting in the checkout
    $pull = git pull --rebase --autostash -X theirs 2>&1
    if($LASTEXITCODE -ne 0){ git rebase --abort 2>&1 | Out-Null; throw "git pull failed: $pull" }
    $push = git push 2>&1
    if($LASTEXITCODE -ne 0){ throw "git push failed: $push" }
  } finally { $ErrorActionPreference=$prevEap; Pop-Location }
  Log "Pushed. Done."

  # ======================= DISPATCH (weighbridge) =======================
  # Separate repo, fully NON-FATAL: a dispatch failure logs a warning and leaves
  # the stock refresh above (already pushed) completely intact. data.csv simply
  # keeps its previous contents until the next 15-min run succeeds.
  try{
    if(-not (Test-Path (Join-Path $DispatchRepo '.git'))){
      Log "Dispatch repo missing at $DispatchRepo - cloning $DispatchRepoUrl ..."
      $prevEap=$ErrorActionPreference; $ErrorActionPreference='Continue'
      git clone $DispatchRepoUrl $DispatchRepo 2>&1 | Out-Null
      $cloneRc=$LASTEXITCODE; $ErrorActionPreference=$prevEap
      if($cloneRc -ne 0){ throw "clone failed (rc=$cloneRc) - point `$DispatchRepo at an existing mis-portal clone" }
    }
    Log "Querying weighbridge (dispatch)..."
    $wb = Invoke-Sql $wbSql
    Log "  $($wb.Rows.Count) weighbridge rows"
    if(@($wb.Rows).Count -eq 0){ throw "weighbridge query returned 0 rows - not overwriting data.csv" }

    $DispatchDir = Join-Path $DispatchRepo 'dispatch'
    if(-not (Test-Path $DispatchDir)){ New-Item $DispatchDir -ItemType Directory -Force | Out-Null }
    [IO.File]::WriteAllText((Join-Path $DispatchDir 'data.csv'), (ConvertTo-CsvText $wb), $enc)
    Log "dispatch\data.csv written ($($wb.Rows.Count) rows)."

    Push-Location $DispatchRepo
    $prevEap=$ErrorActionPreference; $ErrorActionPreference='Continue'
    try{
      git add dispatch/data.csv 2>&1 | Out-Null
      $stamp=[DateTime]::Now.ToString('yyyy-MM-dd HH:mm')
      git commit -m "dispatch refresh $stamp" 2>&1 | Out-Null       # no-op if nothing changed
      $pull = git pull --rebase --autostash -X theirs 2>&1
      if($LASTEXITCODE -ne 0){ git rebase --abort 2>&1 | Out-Null; throw "git pull failed: $pull" }
      $push = git push 2>&1
      if($LASTEXITCODE -ne 0){ throw "git push failed: $push" }
    } finally { $ErrorActionPreference=$prevEap; Pop-Location }
    Log "Dispatch pushed. Done."
  }
  catch{ Log "WARN: dispatch refresh failed ($($_.Exception.Message)) - stock is unaffected; data.csv keeps its previous contents" }
}
catch{
  Log "ERROR: $($_.Exception.Message)"
  exit 1
}
finally{
  if(Test-Path $LockPath){ Remove-Item $LockPath -Force -ErrorAction SilentlyContinue }
}
