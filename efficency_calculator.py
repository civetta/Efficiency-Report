import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import colors
from openpyxl.styles import Fill
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import numpy as np
import re
wb = load_workbook(filename = 'LeadBook.xlsx',data_only = True)


def find_blocks(ws,day):
    max_col = ws.max_column
    max_row = ws.max_row
    col = 3
    ws.cell(row=1,column=max_col+7,value =day+" Summary")
    while col <= max_col:
        teacher_name=ws.cell(row=1,column=col).value
        row = 2
        create_tables(teacher_name,col,max_col,ws,day)
        time_block, tab_block, cell_block = [], [], []
        block_on=False
        block_count=0
        while row < max_row:
            current_cell = ws.cell(row = row,  column = col)
            if current_cell.font.bold==True:
                block_on=True
                time=ws.cell(row = row, column = 1).value
                time=time[8:]
                time_block.append(time)
                tab_block.append(ws.cell(row = row, column = 2).value)
                cell_block.append(current_cell.value)
            if current_cell.font.bold==False:
                if block_on==True:
                    block_count=block_count+1
                    block_tables(col,max_col,time_block,tab_block,cell_block,block_count,ws,teacher_name,day)
                    time_block, tab_block, cell_block = [], [], []
                    block_on=False
            row=row+1  
        col=col+1
                  

def create_tables(teacher_name,col,max_column,ws,day):
    col=col-2
    s_row=(col*8)-7
    ws.cell(row=s_row,column=max_column+2, value = teacher_name)
    ws.cell(row=s_row,column=max_column+3, value = "Average Students")
    ws.cell(row=s_row,column=max_column+4, value ="Average Tabby")
    ws.cell(row=s_row,column=max_column+5, value ="Students/Tabby (Efficiency Score)")

    ws.cell(row=2,column=max_column+7+col, value=teacher_name)
    ws.cell(row=2,column=max_column+7, value ="Name")
    ws.cell(row=3,column=max_column+7, value ="Daily Average Efficiency Score")
    ws.cell(row=4,column=max_column+7, value ="Block 1")
    ws.cell(row=5,column=max_column+7, value ="Block 2")
    ws.cell(row=6,column=max_column+7, value ="Block 3")
    ws.cell(row=7,column=max_column+7, value ="Block 4")
    ws.cell(row=8,column=max_column+7, value ="Block 5")
    ws.cell(row=9,column=max_column+7, value ="Block 6")

def block_tables(col,max_column,time_block,tab_block,cell_block,block_count,ws,teacher_name,day):
    col=col-2
    s_row=(col*8)-7
    ws.cell(row=s_row+block_count,column=max_column+2, value =time_block[0]+" - "+time_block[-1])
    ws.cell(row=s_row+block_count,column=max_column+3, value =round(np.average(cell_block),2))
    tab_block=np.array(tab_block).astype(np.float)
    if sum(tab_block)>2:
        ws.cell(row=s_row+block_count,column=max_column+4, value =round(sum(tab_block)/len(tab_block)))
    
        ws.cell(row=s_row+block_count,column=max_column+5, value =round(round(np.average(cell_block),2)/round(np.average(tab_block),2),2))

        ws.cell(row=block_count+3,column=max_column+7+col,value =round(round(np.average(cell_block),2)/round(np.average(tab_block),2),2))   

def create_block_table(wb,lead_name):
    week=["Mon", "Tue", "Wed","Thu","Fri","Sat","Sun"]
    for day in week:
        ws = wb.get_sheet_by_name(day)
        max_column = ws.max_column
        max_col=find_blocks(ws,day)
        wb.save(lead_name[0:3]+'.xlsx')
    return max_column
