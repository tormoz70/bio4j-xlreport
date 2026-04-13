@echo off
REM ==========================================================================
REM  setup_macros.cmd — Copy Module1.js to the legacy reports directory.
REM  Run this ONCE before using any run_as_console_macro_*.cmd script.
REM ==========================================================================
setlocal

set SCRIPT_DIR=%~dp0
set EXAMPLES_DIR=%SCRIPT_DIR%examples
set RPTS_DIR=c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_olds\01_givc

if not exist "%EXAMPLES_DIR%\Module1.js" (
    echo ERROR: %EXAMPLES_DIR%\Module1.js not found.
    exit /b 1
)

echo Copying Module1.js to %RPTS_DIR% ...
copy /Y "%EXAMPLES_DIR%\Module1.js" "%RPTS_DIR%\Module1.js"
if errorlevel 1 (
    echo ERROR: Copy failed.
    exit /b 1
)
echo Done. Module1.js is ready for macro dispatch.
endlocal
