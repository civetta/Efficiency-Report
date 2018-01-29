import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import colors
from openpyxl.styles import Fill
from openpyxl.styles import PatternFill
import numpy as np
import re

def find_average(ws,max_col,col):
    while col <= max_col+1:
        row=1
        num_list=[]
        while row<9:
            o_cell = ws.cell(row=row, column=col)
            if o_cell.value != None:
                try:
                    value = int(o_cell.value)
                    num_list.append(o_cell.value)
                except ValueError:
                    pass
            row=row+1
        if len(num_list)>0:
            ws.cell(row=3,column=col,value = round(sum(num_list)/len(num_list),2))
        col=col+1

def create_summary(wb,start_col):
    week=["Mon", "Tue", "Wed","Thu","Fri","Sat","Sun"]
    weekly_breakdown = wb.create_sheet("Weekly Breakdown",0)
    day_row=1
    for day in week:
        ws=wb.get_sheet_by_name(day)
        max_col=ws.max_column
        col=start_col+8
        find_average(ws,max_col,col)
        while col <= max_col+1:
            new_row=day_row
            row=1
            while row<9:
                o_cell = ws.cell(row=row, column=col-1)
                n_cell = weekly_breakdown.cell(row=new_row,column=col-7-start_col)
                n_cell.value=o_cell.value
                row=row+1
                new_row=new_row+1
            col=col+1
        day_row=new_row+3



