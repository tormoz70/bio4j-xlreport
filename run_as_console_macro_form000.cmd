@echo off
REM ==========================================================================
REM  form000 — Свод кинопоказа (Module1.m1)
REM
REM  Macro: macroAfter name="Module1.m1"
REM  After report is built, m1() creates an aggregated summary sheet "Свод".
REM
REM  Prerequisite: run run_as_console_macro_setup.cmd once to deploy Module1.js
REM
REM  SQL bind params: date_from, date_to, org_id, pu_number,
REM                   SYS_CURUSERROLES, SYS_CURODEPUID, reliable
REM ==========================================================================
setlocal

if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM ── Report files ────────────────────────────────────────────────────────────
set RPT_DIR=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc
set RPT_XML=%RPT_DIR%\form000(rpt).xml
set TEMPLATE_XLSX=%RPT_DIR%\form000(rpt).xlsm
set OUT_FILE=c:\data\prjs\bio4j-xlreport\out\form000-macro-m1.xlsm

REM ── Load credentials from env_oracle.cmd if env vars are not set ────────────
if "%ORACLE_DB_USER%"=="" if exist "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd" call "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd"
if "%ORACLE_DB_USER%"=="" ( echo ERROR: ORACLE_DB_USER not set. Copy env_oracle.sample.cmd to env_oracle.cmd and fill in credentials. & exit /b 1 )

set DB_URL=%ORACLE_DB_URL%
set DB_USER=%ORACLE_DB_USER%
set DB_PASSWORD=%ORACLE_DB_PASSWORD%
set DB_DRIVER=%ORACLE_DB_DRIVER%
set DB_FETCH_SIZE=%ORACLE_DB_FETCH_SIZE%

REM ── Mode: lenient because template is xlsm (liveScripts flag present) ──────
set MODE=lenient

REM ── Runtime parameters ──────────────────────────────────────────────────────
REM  date_from / date_to  — reporting period
REM  org_id               — organisation ID
REM  pu_number            — film registration number
REM  reliable             — 1 = exclude bad sessions, 0 = all
set RPT_PRMS=date_from=2026-03-01;date_to=2026-03-31;org_id=;pu_number=;reliable=1

REM  User context
set RPT_USER_UID=%ORACLE_RPT_USER_UID%
set RPT_USER_ORG=%ORACLE_RPT_USER_ORG%
set RPT_USER_ROLES=%ORACLE_RPT_USER_ROLES%

set STOP_ON_FINISH=true

echo Running form000 (Module1.m1) from Oracle...
mkdir c:\data\prjs\bio4j-xlreport\out 2>nul
call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"

endlocal
