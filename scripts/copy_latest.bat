@echo off
set SOURCE="C:\Users\I349481\OneDrive - SAP SE\Documents - BT&A EA APAC\Regional Folder\1) EA-RISE General\17_Performance Goal\AFO_EA_Impact_KPI___cACV_Territory_Analyzer_Ops_updated.xlsx"
set DEST="C:\Users\I349481\Downloads\EA-KPI-Dashboard\data\latest.xlsx"

if not exist %SOURCE% (
    echo ERROR: Source file not found: %SOURCE%
    exit /b 1
)

copy /Y %SOURCE% %DEST%
echo Copied to %DEST% on %date% %time%
