# users_apps
This is a powershell program I wrote, or re-wrote really, for a client when the previous version broke. 

## Overview
This program is very simple: 
  - it reads in a servers.csv
  - reads in an applications.csv
  - connects to all of the servers in parallel
    - checks to see if any of the applications listed in applications.csv are running
  - creates a csv (and optionaly an excel file) with a log of the users, what they are running, and how long it has been running for
  
## Run as follows: 


     powershell.exe -executionpolicy unrestricted -command  .\users_apps.ps1  -Servers '.\servers.csv' -Apps '.\applications.csv' -Output 'output.csv' -Excel

