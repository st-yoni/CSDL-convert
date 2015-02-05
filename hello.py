# -*- coding: utf-8 -*-
"""
Created on Wed Jan 28 22:40:03 2015

@author: st-yoni
"""
import xlrd
import xlsxwriter
from Tkinter import Tk
from tkFileDialog import askopenfilename 
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # print number of sheets
    print book.nsheets
 
    # print sheet names
    print book.sheet_names()
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
 
    # read a row
    print first_sheet.row_values(0)
 
    # read a cell
    cell = first_sheet.cell(0,0)
    print cell
    print cell.value
 
    # read a row slice
    print first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2)

def writeSingleCol(path,worksheet,textToWrite):
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet(worksheet)
    index = 0
    for cell in textToWrite:
        worksheet.write(index,0,cell)
        index =index +1
    workbook.close()

    
def writeSingleCell(path,worksheet,textToWrite):
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet(worksheet)
    toWrite = 'twitter.user.name contains_all "'
    for cell in textToWrite:
        toWrite = toWrite + cell.value + ","
    toWrite = toWrite + '"'
    worksheet.write(0,0,toWrite)
    workbook.close()

def readAllCol(path,sheetName):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name(sheetName)
    FullCol = []
    for cell in sheet.col(0):
        FullCol.append(cell)
    return FullCol
    

#----------------------------------------------------------------------
if __name__ == "__main__":
    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    path = askopenfilename() # show an "Open" dialog box and return the path to the selected file  
    writeSingleCol(path,"source",["test","yoni","steinmetz","Bunz"])
    print readAllCol(path,"source")
    writeSingleCell(path,"dest",readAllCol(path,"source"))
    print readAllCol(path,"dest")
    