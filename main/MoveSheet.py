'''
Created on 26 avr. 2013

@author: Alexis Thongvan

The purpose of this code is to move a sheet across the
existing one
'''
from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def MoveSheet(sheet, new_position, wb):
    """sheet can be an integer (the number of the sheet)
    or a String (the name of the sheet it will be move
    next to)"""
    wb.Sheets(sheet).Move(Before=wb.Sheets(new_position))

MoveSheet("toto", "lol", wb)

saveCopy("../files/testWorkbookMoveSheetResult.xlsx", wb)
closeExcel(xl)
#MoveSheet("toto", 3, wb) produce the same result

#Before : "toto", "titi", "lol"
#After : "titi", "toto", "lol"
