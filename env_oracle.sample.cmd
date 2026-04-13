@echo off
REM ==========================================================================
REM  env_oracle.sample.cmd — Template for Oracle credentials.
REM
REM  HOW TO USE:
REM    1. Copy this file to %USERPROFILE%\.bio4j-xlreport\env_oracle.cmd
REM         mkdir "%USERPROFILE%\.bio4j-xlreport"
REM         copy env_oracle.sample.cmd "%USERPROFILE%\.bio4j-xlreport\env_oracle.cmd"
REM    2. Open the copy and fill in the actual values.
REM    3. The file lives OUTSIDE the repository — never commit it.
REM ==========================================================================

REM ── Oracle JDBC connection ───────────────────────────────────────────────────
set ORACLE_DB_URL=jdbc:oracle:thin:@<host>:<port>:<sid>
set ORACLE_DB_USER=<username>
set ORACLE_DB_PASSWORD=<password>
set ORACLE_DB_DRIVER=oracle.jdbc.OracleDriver
set ORACLE_DB_FETCH_SIZE=1000

REM ── Default report user context ──────────────────────────────────────────────
set ORACLE_RPT_USER_UID=<user-uid-guid>
set ORACLE_RPT_USER_ORG=<org-id>
set ORACLE_RPT_USER_ROLES=1,6
