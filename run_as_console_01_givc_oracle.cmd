@echo off
setlocal

REM Optional Oracle client setup (if needed by your environment)
if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM -----------------------------
REM Configure report input/output
REM -----------------------------
set RPT_XML=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc\form004(rpt).xml
set TEMPLATE_XLSX=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc\form004(rpt).xlsx
set OUT_FILE=c:\data\prjs\bio4j-xlreport\out\form004-result-oracle.xlsx

REM ── Load credentials from env_oracle.cmd if env vars are not set ────────────
if "%ORACLE_DB_USER%"=="" if exist "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd" call "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd"
if "%ORACLE_DB_USER%"=="" ( echo ERROR: ORACLE_DB_USER not set. Copy env_oracle.sample.cmd to env_oracle.cmd and fill in credentials. & exit /b 1 )

set DB_URL=%ORACLE_DB_URL%
set DB_USER=%ORACLE_DB_USER%
set DB_PASSWORD=%ORACLE_DB_PASSWORD%
set DB_DRIVER=%ORACLE_DB_DRIVER%
set DB_FETCH_SIZE=%ORACLE_DB_FETCH_SIZE%

REM compatibility: strict | lenient
set MODE=lenient

REM Runtime report parameters used by SQL binds (:date_from, :date_to)
set RPT_PRMS=date_from=2026-03-21;date_to=2026-03-31;holding_id=70;org_id=7667;pu_number=111005326
set RPT_USER_UID=%ORACLE_RPT_USER_UID%
set RPT_USER_ORG=%ORACLE_RPT_USER_ORG%
set RPT_USER_ROLES=%ORACLE_RPT_USER_ROLES%
set STOP_ON_FINISH=true

echo Running console report builder from Oracle...
call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"

endlocal
