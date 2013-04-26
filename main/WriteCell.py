'''
Created on 26 avr. 2013

@author: Alexis Thongvan

Write data in the cell of a given sheet
the sheet do not need to be visible or selected
'''

from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def WriteCellWithCoordAsNumber(sheet, line, column, value, wb):
    """write a cell using number to find the required
    cell, note that line and column start at 1"""
    wb.Sheets(sheet).Cells(line, column).Value = value

WriteCellWithCoordAsNumber("toto", 1, 1, "I JUST WROTE SOMETHING", wb)

saveCopy("../files/testWorkbookWriteCellResult.xlsx", wb)
closeExcel(xl)
