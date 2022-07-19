@echo off
powershell.exe -executionpolicy unrestricted -command  .\users_apps.ps1  -Servers '.\servers.csv' -Apps '.\applications.csv' -Output 'output.csv' -Excel
::
::powershell.exe -executionpolicy unrestricted -command  C:\Scripts\UpWork\users_apps\users_apps.ps1 -Append -Servers 'C:\Scripts\UpWork\users_apps\servers-to-scan.csv' -Apps 'C:\Scripts\UpWork\users_apps\applications-to-scan.csv' -Output '\\CTO-APPS17A\TandemDrive\_ImportDataCenterStatistics\users_apps\users_apps.csv'