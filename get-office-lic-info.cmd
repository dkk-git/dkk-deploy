@echo off

REM vytvoření adresáře C:\Temp
if not exist C:\Temp (
    mkdir C:\Temp
)

set OUT=C:\Temp\office_lic.txt

set OSPP1="C:\Program Files\Microsoft Office\Office16\OSPP.VBS"
set OSPP2="C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"

if exist %OSPP1% (
    cscript //nologo %OSPP1% /dstatus > "%OUT%" 2>&1
    exit /b 0
)

if exist %OSPP2% (
    cscript //nologo %OSPP2% /dstatus > "%OUT%" 2>&1
    exit /b 0
)

echo OSPP.vbs nebyl nalezen > "%OUT%"
exit /b 1
