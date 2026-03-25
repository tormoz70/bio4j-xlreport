@echo off
setlocal

if not "%ORACLE_HOME%"=="" set "PATH=%ORACLE_HOME%\bin;%PATH%"

call gradlew.bat :app-console:run --args="console"

endlocal
