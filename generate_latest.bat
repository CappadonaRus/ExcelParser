@echo off
set /p fileName="Please, enter file name:"
call java -classpath ExcelParser.jar ru.excel.late_report.LateReportMain %fileName%
pause