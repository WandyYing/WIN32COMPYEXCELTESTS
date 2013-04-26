'''
Created on 26 avr. 2013

@author: Alexis Thongvan

Contains various method to open, save, etc
'''

from win32com import client
import os

def openExcel():
    xl = client.Dispatch("Excel.Application")
    return xl

def openWorkbook(workbook_with_path, xl):
    #you can't use a relative path, because excel won't know it's relative
    #to what, note that I'm using "/" on windows and not "\" as path separators
    wb = xl.Workbooks.Open(os.path.abspath(workbook_with_path))
    xl.Visible = 1
    return wb

def save(wb):
    wb.Save()

def saveCopy(new_name_with_path, wb):
    wb.SaveAs(os.path.abspath(new_name_with_path))

def closeExcel(xl):
    xl.Quit()
