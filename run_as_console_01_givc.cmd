@echo off
setlocal

REM Optional Oracle client setup (if needed by your environment)
if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

REM -----------------------------
REM Configure report input/output
REM -----------------------------
set RPT_XML=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_v2\01_givc\form100(rpt).xml
set TEMPLATE_XLSX=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_v2\01_givc\form100(rpt).xlsm
set DATA_JSON=c:\data\prjs\bio4j-xlreport\examples\data-form100.json
set OUT_XLSX=c:\data\prjs\bio4j-xlreport\out\form100-result.xlsx

REM compatibility: strict | lenient
set MODE=lenient

REM Optional legacy-style runtime parameters
set RPT_PRMS=date_from=2026-01-01;date_to=2026-01-31
set RPT_USER_UID=localuser
set RPT_USER_ORG=5567
set RPT_USER_ROLES=*
set STOP_ON_FINISH=true

echo Running console report builder...
call gradlew.bat :app-console:run --args="/rpt:%RPT_XML% /template:%TEMPLATE_XLSX% /data:%DATA_JSON% /out:%OUT_XLSX% /mode:%MODE% /rptPrms:%RPT_PRMS% /rptUserUID:%RPT_USER_UID% /rptUserOrgId:%RPT_USER_ORG% /rptUserRoles:%RPT_USER_ROLES% /rptStopOnFinish:%STOP_ON_FINISH%"

endlocal
