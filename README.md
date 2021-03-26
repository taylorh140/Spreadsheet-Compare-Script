# Spreadsheet-Compare-Script
A command line interface for calling Spreadsheetcompare.exe included with MS office (some versions)

## details

This script allows calls to the spreadsheetcompare.exe as far as i know the location of the exe is:

*"C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE"*

This is an important line in the script and if the path is wrong it will not run. It would be good to validate this.

from what I can tell the program will compare two xlsx files from a command if the two file paths are placed into a text file with the left file being in line 1 and the right file being in line 2. 
It is also important that the paths **do not have quotes around them** this can be easily avoided in the batch script by using the tilda operator.

## Usage

SpreadsheetCompare.bat C:\file1.xlsx C:\file2.xlsx
