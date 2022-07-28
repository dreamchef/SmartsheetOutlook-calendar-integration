@echo off
set /p id=Enter Smartsheet Access Token: 

echo %id% > assets/SmartsheetAccessToken.txt

set /p start=Enter Start Date (MM/DD/YYYY): 

echo %start% >> assets/SmartsheetAccessToken.txt

set /p end=Enter End Date (MM/DD/YYYY): 

echo %end% >> assets/SmartsheetAccessToken.txt


pause