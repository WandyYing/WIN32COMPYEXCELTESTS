'''
Created on 8 mai 2013

@author: thongvan

Rename a sheet
'''
from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def renameSheet(old_sheet, new_name, wb):
    """Rename a sheet to a new name, old
    sheet can be the number of the sheet, or
    it's actual name as a string"""
    ws = wb.Sheets(old_sheet)
    ws.Name = new_name

renameSheet("lol", "NO", wb)
renameSheet(1, "one", wb)

saveCopy("../files/testWorkbookRenameSheetResult.xlsx", wb)
closeExcel(xl)

#Before : "toto", "titi", "lol"
#After : "one", "titi", "NO"
