'''
Created on 8 mai 2013

@author: thongvan

Rename a sheet
'''
from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def DeleteSheet(sheet, wb, xl):
    """Delete a sheet,
    sheet can be the number of the sheet, or
    it's actual name as a string"""
    xl.displayAlerts = False
    ws = wb.Sheets(sheet)
    ws.Delete()
    xl.displayAlerts = True


DeleteSheet("lol", wb, xl)
DeleteSheet(1, wb, xl)

saveCopy("../files/testWorkbookDeleteSheetResult.xlsx", wb)
closeExcel(xl)

#Before : "toto", "titi", "lol"
#After : NO"
