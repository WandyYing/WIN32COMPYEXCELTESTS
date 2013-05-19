'''
Created on 26 avr. 2013

@author: Alexis Thongvan

Some sheets can be hidden, masked, to display them
right clic on any sheet name in excel and choose display.
then select one of the hidden sheets.

Our testWorkbook.xls contains 2 hidden sheet :
Hidden1 and data
'''
from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def displayHiddenSheet(sheet_name, wb):
    wb.Sheets(sheet_name).Visible = True

displayHiddenSheet("Hidden1", wb)
saveCopy("../files/testWorkbookDisplayHiddenSheetResult.xlsx", wb)
closeExcel(xl)

#As you can see the hidden sheet is now selectable and visible
#Before : "toto", "titi", "lol"
#After : "toto", "titi", "lol", "Hidden1"
