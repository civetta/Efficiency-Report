import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import colors
from openpyxl.styles import Fill
from openpyxl.styles import PatternFill
import numpy as np
import re

def find_row(wb):
    breakdown = wb.get_sheet_by_name("Weekly Breakdown")
    max_row=breakdown.max_row
    summary=create_summary(wb,breakdown)
    current_row=1
    past_row=3
    while current_row<max_row:
        current_cell=breakdown.cell(row=current_row, column=1).value
        if current_cell =="Daily Average Efficiency Score":
            copy_row(past_row,breakdown,summary,current_row)
            summary.cell(row=past_row, column=1, value=breakdown.cell(row=current_row-2, column=1).value)
            past_row=past_row+1
        current_row=current_row+1
    find_average(summary)
    wb.save('LeadBook.xlsx')


def copy_row(past_row,breakdown,summary,row):
    i=1
    while i <= breakdown.max_column:
        summary.cell(row=past_row, column=i, value=breakdown.cell(row=row, column=i).value)
        i=i+1

def create_summary(wb,breakdown):
    summary = wb.create_sheet("Summary", 0)
    summary.cell(row=1, column=1, value="Weekly Summary")
    summary.cell(row=10, column=1, value="Weekly Average")
    i=1
    while i <=breakdown.max_column:
        current_cell = breakdown.cell(row=2, column=i).value
        summary.cell(row=2, column=i,value=breakdown.cell(row=2, column=i).value)
        i=i+1
    return summary
                
def find_average(summary):
    max_col=summary.max_column
    col=2
    while col<=max_col:
        average_list=[]
        row=3
        while row <= 9:
            celler=summary.cell(row=row,column=col).value
            if celler != None:
                try:
                    value = int(celler)
                    average_list.append(celler)
                except ValueError:
                    pass
            row=row+1
        if len(average_list)>1:
            av=(sum(average_list)/len(average_list))
            summary.cell(row=10,column=col, value=round(av,2))
        if len(average_list)is 1:
            summary.cell(row=10,column=col, value=average_list[0])
        col=col+1
            

