import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import colors
from openpyxl.styles import Fill
from openpyxl.styles import PatternFill
import numpy as np
import re
from openpyxl.utils import get_column_letter


def first_two_columns(old,new):
    for row in range(1,old.max_row+1):
        for col in range(1,3):
                val=old.cell(row=row,column=col).value
                try:
                    val=int(val)
                except:
                    val=str(val)
                new.cell(row=row,column=col,value=val)

def first_row(old,new):
    for col in range(1,old.max_column+1):
        new.cell(row=1,column=col,value=old.cell(row=1,column=col).value)

def find_difference(old,new):
    for row in range(3,old.max_row+1):
        for col in range(3,old.max_column+1):
            cell2=(old.cell(row=row,column=col).value)
            cell1=(old.cell(row=row-1,column=col).value)
            try:
                difference=int(cell2)-int(cell1)
            except:
                difference=None
            new.cell(row=row,column=col,value=difference)

            
def make_sheet(wb):
    old = wb.get_sheet_by_name("Raw Pulls")
    new = wb.create_sheet('Raw Changes')    
    first_two_columns(old,new)
    first_row(old,new)
    find_difference(old,new)
    old.column_dimensions[get_column_letter(1)].width  =  int(25)
    new.column_dimensions[get_column_letter(1)].width  =  int(25)
    return wb



