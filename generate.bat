@echo off
set /p fileName="Please, enter file name:"
call java -classpath SQLXmlGenerator.jar ru.vtb.SqlToXmlGenerator.XmlGenerator %fileName%
pause