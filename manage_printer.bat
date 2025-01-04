@echo off
:: Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :run
) else (
    echo Requesting administrative privileges...
    goto :elevate
)

:elevate
powershell -Command "Start-Process cmd -ArgumentList '/c cd /d %~dp0 && python src\manage_printer.py %*' -Verb RunAs"
goto :eof

:run
python src\manage_printer.py %* 