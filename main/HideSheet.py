'''
Created on 26 avr. 2013

@author: Alexis Thongvan

Some sheets can be hidden, masked, to display them
right clic on any sheet name in excel and choose display.
then select one of the hidden sheets.

Our testWorkbook.xls contains 2 hidden sheet :
Hidden1 and data

This piece of code will hide the "titi" sheet
'''

from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def HideSheet(sheet, wb):
    wb.Sheets(sheet).Visible = False

HideSheet("titi", wb)

saveCopy("../files/testWorkbookHideSheetResult.xlsx", wb)
closeExcel(xl)

#As you can see the "titi" sheet is now hidden
#Before : "toto", "titi", "lol"
#After : "toto, "lol"

#Note : HideSheet(2, wb) would have produced the same result
