'''
Created on 26 avr. 2013

@author: Alexis Thongvan

The purpose of this code is to be able to insert a new
sheet at a wanted position.
'''

from Basics import openExcel, openWorkbook, saveCopy, closeExcel

xl = openExcel()
wb = openWorkbook("../files/testWorkbook.xlsx", xl)

def insertSheetAt(position, name, workbook):
    """Position can be an integer (position of the sheet), or
    its name (String).
    name is the name of the new sheet (DUH)."""

    #Selecting the position, the sheet will be inserted
    #one step before.
    workbook.Sheets(position).Select()
    #Inserting a blank sheet.
    new_ws = wb.Worksheets.Add()
    #Renaming.
    new_ws.Name = name

insertSheetAt(2, "firstNewSheet", wb)
insertSheetAt("lol", "SecondNewSheet", wb)

saveCopy("../files/testWorkbookInsertNewSheetResult.xlsx", wb)
closeExcel(xl)

#Voila.
#At first you add 3 sheets : "toto", "titi", "lol"
#After the execution you then have
#"toto", "firstNewSheet",  "titi", "SecondNewSheet", "lol"
