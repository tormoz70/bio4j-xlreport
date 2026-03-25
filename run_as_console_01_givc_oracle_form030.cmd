@echo off
setlocal

REM Optional Oracle client setup (if needed by your environment)
if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM -----------------------------
REM Configure report input/output
REM -----------------------------
set RPT_XML=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc\form030(rpt).xml
set TEMPLATE_XLSX=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc\form030(rpt).xlsx
set OUT_FILE=c:\data\prjs\bio4j-xlreport\out\form030-result-oracle.xlsx

REM -----------------------------
REM Oracle JDBC connection settings
REM -----------------------------
set DB_URL=jdbc:oracle:thin:@192.168.70.29:1521:GIVCDB
set DB_USER=GIVCADMIN
set DB_PASSWORD=j12
set DB_DRIVER=oracle.jdbc.OracleDriver
set DB_FETCH_SIZE=1000

REM compatibility: strict | lenient
set MODE=lenient

REM Runtime report parameters used by SQL binds
set RPT_PRMS=date_from=2026-03-21;date_to=2026-03-31;anotverified=0
set RPT_USER_UID=38522A3CD08D437DE0531E32A8C08E11
set RPT_USER_ORG=5567
set RPT_USER_ROLES=1,6
set STOP_ON_FINISH=true

echo Running console report builder from Oracle (form030)...
call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /out:%OUT_FILE% /mode:%MODE% /dbUrl:%DB_URL% /dbUser:%DB_USER% /dbPassword:%DB_PASSWORD% /dbDriver:%DB_DRIVER% /dbFetchSize:%DB_FETCH_SIZE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"

endlocal
