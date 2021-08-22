@echo off
set /p fileName="Please, enter file name:"
call java -classpath ExcelParser.jar ru.excel.ReportCreator %fileName%
pause