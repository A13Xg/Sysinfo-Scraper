@echo off
REM This script is designed to grab the current system information and dump it to a TXT file.
REM Signed by @13X

echo Current User: %USERNAME% >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

echo Computer Name: %COMPUTERNAME% >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

echo Current IP Configuration: >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
ipconfig /all >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

echo System Info: >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
Systeminfo >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

