@echo off
REM This script is designed to grab the current system information and dump it to a TXT file.
REM Signed by @13X

REM '%USERNAME%' is a variable that grabs the currently logged in username and makes it a string.
REM '%COMPUTERNAME%' is a variable that grabs the computers name such as DESKTOP-ABC123.


REM Grab the current user and computer name.
echo Current User: %USERNAME% >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
echo Computer Name: %COMPUTERNAME% >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

REM WMIC command to grab the unit's SerialNumber from BIOS.
WMIC BIOS get SerialNumber >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

REM Seperator before listing system info.
echo. >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
echo ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
echo. >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

REM Title 'System Info:'
echo System Info: >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
REM Generate system info and send to text file.
Systeminfo >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

REM Seperator before listing IP Configuration.
echo. >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
echo. ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
echo. >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT

REM List IP configuration and send to text file.
echo Current IP Configuration: >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
ipconfig /all >> sysinfo_%USERNAME%_%COMPUTERNAME%.TXT
