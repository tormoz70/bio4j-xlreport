@echo off
REM ==========================================================================
REM  form166 — Сводка по регионам (Module1.Macros1)
REM
REM  Macro: macroAfter name="Module1.Macros1"
REM  Macros1() is currently a stub in Module1.js — implement logic if needed.
REM
REM  Prerequisite: run run_as_console_macro_setup.cmd once to deploy Module1.js
REM
REM  SQL bind params: region_id, month_from, month_to
REM  month_from / month_to format: YYYY-MM  (e.g. 2026-01)
REM
REM  Note: XML has liveScripts="true" — requires MODE=lenient.
REM ==========================================================================
setlocal

if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM ── Report files ────────────────────────────────────────────────────────────
set RPT_DIR=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc
set RPT_XML=%RPT_DIR%\form166(rpt).xml
set TEMPLATE_XLSX=%RPT_DIR%\form166(rpt).xlsm
set OUT_FILE=c:\data\prjs\bio4j-xlreport\out\form166-macro-Macros1.xlsm

REM ── Oracle JDBC ─────────────────────────────────────────────────────────────
set DB_URL=jdbc:oracle:thin:@192.168.70.29:1521:GIVCDB
set DB_USER=GIVCADMIN
set DB_PASSWORD=j12
set DB_DRIVER=oracle.jdbc.OracleDriver
set DB_FETCH_SIZE=1000

REM ── Mode: lenient required (liveScripts="true") ─────────────────────────────
set MODE=lenient

REM ── Runtime parameters ──────────────────────────────────────────────────────
REM  region_id  — region UID from reg$tkladr_regs
REM  month_from — start month in YYYY-MM format
REM  month_to   — end month in YYYY-MM format
set RPT_PRMS=region_id=;month_from=2026-01;month_to=2026-03

REM  User context
set RPT_USER_UID=38522A3CD08D437DE0531E32A8C08E11
set RPT_USER_ORG=5567
set RPT_USER_ROLES=1,6

set STOP_ON_FINISH=true

echo Running form166 (Module1.Macros1) from Oracle...
mkdir c:\data\prjs\bio4j-xlreport\out 2>nul
call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"

endlocal
