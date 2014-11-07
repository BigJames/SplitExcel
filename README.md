SplitExcel
==========

Split an Excel File into multiple files based on the worksheets.  This is provided with no license and no guarantee for support.
The executable is run from the command line as such: C:\SplitExcel.exe "someExcelFile.xlsx"

This will take each worksheet in the book and make a separate file for it.  I made this to help with an SSRS report I am building
The report auto generates an Excel File, I merely made a script to run in the SSIS package for the report that will split this
automatically.  It is simple, and does exactly what I said.  I wrote the script in C#.
