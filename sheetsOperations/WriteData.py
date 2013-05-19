'''
Created on 19 mai 2013

@author: Alexis Thongvan

Write data in the row of a given sheet
the sheet do not need to be visible or selected
'''

from workbooksOperations.Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def WriteData(sheet, line, column, data, wb):
    """write a list of list inside an excel sheet,
    line and column are the top left corner of the data."""
    width = len(data[0])
    length = len(data)
    ws = wb.Sheets(sheet)
    ws.Range(ws.Cells(line, column), ws.Cells(line + length - 1, column + width - 1)).Value = data

data = [
        ['a', 'b', 'c', 'd', 'e'],
        ['f', 'g', 'h', 'i', 'j'],
        ['k', 'l', 'm', 'n', 'o'],
        ['p', 'q', 'r', 's', 't'],
        ['u', 'v', 'w', 'x', 'y'],
        ]

WriteData("lol", 15, 6, data, wb)

saveCopy("../files/testWorkbookWriteDataResult.xlsx", wb)
closeExcel(xl)
