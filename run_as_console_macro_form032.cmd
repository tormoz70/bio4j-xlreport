@echo off
REM ==========================================================================
REM  form032 — Сводка по городам/кинотеатрам (Module1.buildAll)
REM
REM  Macro: macroAfter name="Module1.buildAll"
REM  After report is built, buildAll() constructs a city×date×theatre matrix.
REM
REM  Prerequisite: run run_as_console_macro_setup.cmd once to deploy Module1.js
REM
REM  SQL bind params: pu_number (film reg. number), wee_id (weekend period ID)
REM
REM  Note: XML has convertResultToPDF="true" and liveScripts="true" — requires
REM        MODE=lenient to skip unsupported features.
REM ==========================================================================
setlocal

if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM ── Report files ────────────────────────────────────────────────────────────
set RPT_DIR=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc
set RPT_XML=%RPT_DIR%\form032(rpt).xml
set TEMPLATE_XLSX=%RPT_DIR%\form032(rpt).xlsm
set OUT_FILE=c:\data\prjs\bio4j-xlreport\out\form032-macro-buildAll.xlsm

REM ── Load credentials from env_oracle.cmd if env vars are not set ────────────
if "%ORACLE_DB_USER%"=="" if exist "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd" call "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd"
if "%ORACLE_DB_USER%"=="" ( echo ERROR: ORACLE_DB_USER not set. Copy env_oracle.sample.cmd to env_oracle.cmd and fill in credentials. & exit /b 1 )

set DB_URL=%ORACLE_DB_URL%
set DB_USER=%ORACLE_DB_USER%
set DB_PASSWORD=%ORACLE_DB_PASSWORD%
set DB_DRIVER=%ORACLE_DB_DRIVER%
set DB_FETCH_SIZE=%ORACLE_DB_FETCH_SIZE%

REM ── Mode: lenient required (convertResultToPDF + liveScripts flags) ─────────
set MODE=lenient

REM ── Runtime parameters ──────────────────────────────────────────────────────
REM  pu_number — film registration number (e.g. 12345)
REM  wee_id    — weekend period ID from givc_nsi.weekends_per
set RPT_PRMS=pu_number=12345;wee_id=1

REM  User context
set RPT_USER_UID=%ORACLE_RPT_USER_UID%
set RPT_USER_ORG=%ORACLE_RPT_USER_ORG%
set RPT_USER_ROLES=%ORACLE_RPT_USER_ROLES%

set STOP_ON_FINISH=true

echo Running form032 (Module1.buildAll) from Oracle...
mkdir c:\data\prjs\bio4j-xlreport\out 2>nul
call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"

endlocal
