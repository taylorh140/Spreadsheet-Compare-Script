@echo off
rem Spreadsheet compare command for windows.
rem Taylor Hillegeist

echo %~1 >  "%temp%\tmpcmp.txt"
echo %~2 >> "%temp%\tmpcmp.txt"
"C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE" "%temp%\tmpcmp.txt"