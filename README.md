# Spreadsheet-Compare-Script
A command line interface for calling Spreadsheetcompare.exe included with MS office (some versions)

## details

This script allows calls to the spreadsheetcompare.exe the location of the exe is:

*"C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE"*

This is an important line in the script; if the path is wrong it will not run. It would be good to validate this.

The program will compare two xlsx files from a command if the two file paths are placed into a text file with the left file being in line 1 and the right file being in line 2.  That is exactly what the batch script does. After the batch script runs it deletes the temporary text file.
It is also important that the paths **do not have quotes around them** this can be easily avoided in the batch script by using the tilda operator.

## Usage

SpreadsheetCompare.bat C:\file1.xlsx C:\file2.xlsx
