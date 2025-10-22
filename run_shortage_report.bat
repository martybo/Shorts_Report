@echo off
setlocal ENABLEEXTENSIONS

REM =====================================================================
REM   TRUE SHORTAGE REPORT — UNC-PATH SAFE VERSION
REM   Works even if double-clicked directly from \\server\share\folder
REM =====================================================================

REM === Step 1: Map the UNC path to a temporary drive ===
pushd "%~dp0"
if errorlevel 1 (
  echo [ERROR] Could not pushd into "%~dp0"
  echo Try mapping the network share to a drive letter first.
  pause
  exit /b 1
)

echo ===============================================================
echo   TRUE Shortage Report - Run %DATE% %TIME%
echo   Working dir : %CD%
echo ===============================================================
echo.

REM === Step 2: Ensure output folder exists ===
if not exist "output" mkdir "output"

REM === Step 3: Detect Python ===
set "PYEXE="
where py >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE (
  where python >nul 2>&1 && set "PYEXE=python"
)
if not defined PYEXE (
  echo [ERROR] Python not found on PATH.
  pause
  popd
  exit /b 2
)

REM === Step 4: Input files ===
set "ORDERS=Orders Report Generator.csv"
set "PRODUCT=ExportFullProductList.csv"
set "SUBS=Substitutions.csv"
set "SCRIPT=warehouse_shortage_report_v1_nc_frozen.py"

REM === Step 5: Warehouse + NC orderlists ===
set "WH_LIST=Warehouse;Warehouse Controlled Drugs;Warehouse - CD Products"
set "NC_LIST=Supplier;Testers Perfume;Warehouse;Warehouse - CD Products;Xmas Warehouse;Perfumes;AAH (H&B);PHOENIX;YANKEE"

REM === Step 6: Check file presence ===
if not exist "%SCRIPT%" (
  echo [ERROR] Missing script: "%SCRIPT%"
  pause
  popd
  exit /b 3
)
if not exist "%ORDERS%" (
  echo [ERROR] Missing orders file: "%ORDERS%"
  pause
  popd
  exit /b 4
)
if not exist "%PRODUCT%" (
  echo [ERROR] Missing product list: "%PRODUCT%"
  pause
  popd
  exit /b 5
)

set "SUBS_ARG="
if not exist "%SUBS%" (
  echo [WARN] Substitutions file not found ("%SUBS%") - continuing without substitutions.
) else (
  set "SUBS_ARG=--subs "%SUBS%""
)

REM === Step 7: Display parameters ===
echo Using Python: %PYEXE%
echo Script       : %SCRIPT%
echo Orders file  : %ORDERS%
echo Product list : %PRODUCT%
echo Substitutions: %SUBS%
echo Warehouse OLs: %WH_LIST%
echo NC OLs       : %NC_LIST%
echo Output       : output\Shortage_Report.xlsx
echo ===============================================================
echo.

REM === Step 8: Run the script ===
if defined SUBS_ARG (
  %PYEXE% "%SCRIPT%" --orders "%ORDERS%" --product-list "%PRODUCT%" --subs "%SUBS%" --warehouse-orderlists "%WH_LIST%" --nc-orderlists "%NC_LIST%" --out "output\Shortage_Report.xlsx"
) else (
  %PYEXE% "%SCRIPT%" --orders "%ORDERS%" --product-list "%PRODUCT%" --warehouse-orderlists "%WH_LIST%" --nc-orderlists "%NC_LIST%" --out "output\Shortage_Report.xlsx"
)

set "EC=%ERRORLEVEL%"
echo.
echo ExitCode=%EC%
if not "%EC%"=="0" (
  echo [ERROR] Script returned a non-zero exit code.
  echo Check output\run_error_*.log for details.
  pause
  popd
  exit /b %EC%
)

REM === Step 9: Open the newest Excel file in output ===
set "LATEST="
for /f "delims=" %%F in ('dir /b /a:-d /o:-d "output\Shortage_Report*.xlsx" 2^>nul') do (
  set "LATEST=output\%%F"
  goto :FoundLatest
)

echo [INFO] No Excel output found.
goto :End

:FoundLatest
echo Opening "%LATEST%" ...
start "" "%LATEST%"

:End
echo.
pause

REM === Step 10: Unmap temp drive and exit ===
popd
endlocal
exit /b 0
