@echo off
echo Excel File Analyzer
echo ===================
echo This will analyze Excel files and show their structure

if "%~1"=="" (
    echo Usage: analyze_excel.bat [excel_file1] [excel_file2] ...
    echo Example: analyze_excel.bat Clockify_Time_Report.xlsx
    goto :eof
)

echo Installing required packages if needed...
pip install pandas openpyxl

echo Running analysis...
python analyze_excel.py %*

echo Analysis complete.
pause 