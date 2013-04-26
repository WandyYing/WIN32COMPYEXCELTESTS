WIN32COMPYEXCELTESTS
====================

Some test unsing python via win32com to manipulate excel datas

There are 3 Things you need to know :

xl : this is a variable containing the excel application

wb : this is a variable containing the workbook.
A workbook is basically an excel file, a collection of worksheet

ws : this is a variable containing a worksheet.
A worksheet is a sheet from a workbook

_____________________________

File which needs to be installed are in the /installation folder.
Python 2.7 is not enought since we are using some excel constants
(xlUp, xlDown, etc)

All test are being done over a test workbook
located in /files/testWorkbook.xlsx