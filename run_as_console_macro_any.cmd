@echo off
REM ==========================================================================
REM  run_as_console_macro_any.cmd — Run ANY report from rpts_olds/01_givc
REM  by specifying its number.
REM
REM  Usage:
REM    run_as_console_macro_any.cmd <formNumber> [rptPrms]
REM
REM  Examples:
REM    run_as_console_macro_any.cmd 000
REM    run_as_console_macro_any.cmd 000 "date_from=2026-03-01;date_to=2026-03-31;pu_number=12345"
REM    run_as_console_macro_any.cmd 032 "pu_number=12345;wee_id=1"
REM    run_as_console_macro_any.cmd 166 "region_id=;month_from=2026-01;month_to=2026-03"
REM
REM  Prerequisite: run run_as_console_macro_setup.cmd once to deploy Module1.js.
REM ==========================================================================
setlocal

if "%~1"=="" (
    echo Usage: %~nx0 ^<formNumber^> [rptPrms]
    echo   formNumber — 3-digit form number, e.g. 000, 032, 166
    echo   rptPrms    — semicolon-separated key=value pairs
    exit /b 1
)

set FORM_NUM=%~1
if not "%~2"=="" set RPT_PRMS=%~2

if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM ── Report files ────────────────────────────────────────────────────────────
set RPT_DIR=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc
set RPT_XML=%RPT_DIR%\form%FORM_NUM%(rpt).xml
set TEMPLATE_XLSX=%RPT_DIR%\form%FORM_NUM%(rpt).xlsm
set OUT_FILE=c:\data\prjs\bio4j-xlreport\out\form%FORM_NUM%-macro-result.xlsm

if not exist "%RPT_XML%" (
    echo ERROR: Report XML not found: %RPT_XML%
    exit /b 1
)
if not exist "%TEMPLATE_XLSX%" (
    echo ERROR: Template not found: %TEMPLATE_XLSX%
    exit /b 1
)
if not exist "%RPT_DIR%\Module1.js" (
    echo WARNING: Module1.js not found in %RPT_DIR%
    echo          Run run_as_console_macro_setup.cmd first to deploy it.
)

REM ── Load credentials from env_oracle.cmd if env vars are not set ────────────
if "%ORACLE_DB_USER%"=="" if exist "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd" call "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd"
if "%ORACLE_DB_USER%"=="" ( echo ERROR: ORACLE_DB_USER not set. Copy env_oracle.sample.cmd to env_oracle.cmd and fill in credentials. & exit /b 1 )

set DB_URL=%ORACLE_DB_URL%
set DB_USER=%ORACLE_DB_USER%
set DB_PASSWORD=%ORACLE_DB_PASSWORD%
set DB_DRIVER=%ORACLE_DB_DRIVER%
set DB_FETCH_SIZE=%ORACLE_DB_FETCH_SIZE%

REM ── Always lenient — legacy templates have liveScripts/convertResultToPDF ───
set MODE=lenient

REM ── User context ────────────────────────────────────────────────────────────
set RPT_USER_UID=%ORACLE_RPT_USER_UID%
set RPT_USER_ORG=%ORACLE_RPT_USER_ORG%
set RPT_USER_ROLES=%ORACLE_RPT_USER_ROLES%

set STOP_ON_FINISH=true

echo Running form%FORM_NUM% from Oracle (lenient mode, macro dispatch enabled)...
mkdir c:\data\prjs\bio4j-xlreport\out 2>nul

if not "%RPT_PRMS%"=="" (
    call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"
) else (
    call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"
)

endlocal
