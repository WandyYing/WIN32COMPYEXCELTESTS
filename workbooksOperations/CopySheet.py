'''
Created on 26 avr. 2013

@author: Alexis Thongvan

The purpose of this code is to copy one of the
workbook's sheet and rename it
'''

from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def copySheet(source_sheet, new_name, wb):
    """Copy a sheet, source_sheet must be the name in
    a String, you cannot use an integer"""
    ws = wb.Sheets(source_sheet)
    ws.Copy(Before=wb.Sheets(source_sheet))
    ws = wb.Sheets(source_sheet + " (2)")
    ws.Name = new_name

copySheet("lol", "BOOOM", wb)

saveCopy("../files/testWorkbookCopySheetResult.xlsx", wb)
closeExcel(xl)

#The "lol" has been duplicated and rename "BOOOM"
#Before : "toto", "titi", "lol"
#After : "toto", "titi", "BOOOM", "lol"
