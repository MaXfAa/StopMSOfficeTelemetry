@echo off
cls
echo Block Microsoft Office telemetry 
echo Confirm the administrator execution request to pursue.
echo.
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0.\silent-worker.ps1""' -Verb RunAs}"
echo Perfect; the Office telemetry is now blocked! :)
echo.
pause