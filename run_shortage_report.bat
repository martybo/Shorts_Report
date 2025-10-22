@echo off
setlocal ENABLEEXTENSIONS

REM === Base dir (folder of this BAT) ===
set "BASEDIR=%~dp0"
pushd "%BASEDIR%" 1>nul 2>nul

REM === Make sure output folder exists ===
if not exist "output" mkdir "output"

REM === Prefer Python Launcher if available ===
set "PYEXE="
where py >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE (
  where python >nul 2>&1 && set "PYEXE=python"
)
if not defined PYEXE (
  echo [ERROR] Python not found on PATH.
  pause
  exit /b 1
)

REM === Inputs (edit if you use different filenames) ===
set "ORDERS=Orders Report Generator.csv"
set "PRODUCT=ExportFullProductList.csv"
set "SUBS=Substitutions.csv"
set "SUBS_LABEL=%SUBS%"

REM === Script name (edit if you saved with a different filename) ===
set "SCRIPT=warehouse_shortage_report.py"

REM === Warehouse orderlists for Dept/Supplier/Group summaries ===
set "WH_LIST=Warehouse;Warehouse Controlled Drugs;Warehouse - CD Products"

REM === Orderlists to include ONLY for Branch_NC / Company_NC ===
set "NC_LIST=Supplier;Testers Perfume;Warehouse;Warehouse - CD Products;Xmas Warehouse;Perfumes;AAH (H&B);PHOENIX;YANKEE"

echo ===============================================================
echo  TRUE Shortage Report - Run %DATE% %TIME%
echo  CWD  : %CD%
echo  PY   : %PYEXE%
echo  IN   : "%ORDERS%" + "%PRODUCT%" + "%SUBS_LABEL%"
echo  OUT  : output\Shortage_Report_*.xlsx  (final name auto-suffixed with _wcDDMMYY)
echo ===============================================================
echo.

if not exist "%SCRIPT%" (
  echo [ERROR] Missing script: "%SCRIPT%"
  pause
  exit /b 2
)
if not exist "%ORDERS%" (
  echo [ERROR] Missing orders file: "%ORDERS%"
  pause
  exit /b 3
)
if not exist "%PRODUCT%" (
  echo [ERROR] Missing product file: "%PRODUCT%"
  pause
  exit /b 4
)
REM Substitutions are optional; warn if missing
if not exist "%SUBS%" (
  echo [WARN] Substitutions file not found ("%SUBS%") - continuing without substitutions.
  set "SUBS_ARG="
  set "SUBS_LABEL=(none)"
) else (
  set "SUBS_ARG=--subs "%SUBS%""
)

REM === Build command (script handles naming & _wcDDMMYY) ===
set "CMD=%PYEXE% "%SCRIPT%" --orders "%ORDERS%" --product-list "%PRODUCT%""
if defined SUBS_ARG set "CMD=%CMD% %SUBS_ARG%"
set "CMD=%CMD% --warehouse-orderlists "%WH_LIST%" --nc-orderlists "%NC_LIST%" --out "output\Shortage_Report.xlsx""

echo Running:
echo %CMD%
echo.

call %CMD%
set "EC=%ERRORLEVEL%"

echo.
echo ExitCode=%EC%
if not "%EC%"=="0" (
  echo [ERROR] Script returned a non-zero exit code. If an error log was written,
  echo         check "output\run_error_*.log".
  pause
  exit /b %EC%
)

REM === Open the newest Shortage_Report*.xlsx in output ===
for /f "delims=" %%F in ('dir /b /a:-d /o:-d "output\Shortage_Report*.xlsx" 2^>nul') do (
  set "LATEST=output\%%F"
  goto :FoundLatest
)

echo [INFO] No Excel file found in output\ after run.
goto :End

:FoundLatest
echo Opening "%LATEST%" ...
start "" "%LATEST%"

:End
echo.
pause
endlocal
