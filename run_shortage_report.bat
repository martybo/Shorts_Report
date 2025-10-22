@echo off
setlocal

pushd "\\med-vfileprint\Users Shared Folders\Data Analyst\Shorts\NEWEST" || (
  echo Could not access network folder.
  exit /b 1
)

if not exist "output" mkdir "output"

REM Timestamp (robust)
for /F %%T in ('powershell -NoProfile -Command "$([DateTime]::Now.ToString('yyyyMMdd_HHmmss'))"') do set "TS=%%T"

REM Run
set "WAREHOUSE_LIST=Warehouse;Warehouse Controlled Drugs;Warehouse - CD Products"

py -3 warehouse_shortage_report_v1_nc.py ^
  --orders "Orders Report Generator.csv" ^
  --product-list "ExportFullProductList.csv" ^
  --subs "Substitutions.csv" ^
  --warehouse-orderlists "Warehouse;Warehouse Controlled Drugs;Warehouse - CD Products" ^
  --nc-orderlists "Supplier;Testers Perfume;Warehouse;Warehouse - CD Products;Xmas Warehouse;Perfumes;AAH (H&B);PHOENIX;YANKEE" ^
  --out "output\Shortage_Report.xlsx"

echo ExitCode=%ERRORLEVEL%

REM Open newest report (handles wc* rename)
for /f "delims=" %%F in ('dir /b /a:-d /o:-d "output\Shortage_Report*.xlsx"') do (
  set "LATEST=%%F"
  goto :gotlatest
)
:gotlatest
if defined LATEST (
  start "" "output\%LATEST%"
) else (
  echo No Excel file found in output\.
)

pause
