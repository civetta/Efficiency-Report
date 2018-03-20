import openpyxl
from openpyxl import Workbook
from datetime import date
import calendar


def create_books(max_col,wb):
    for teacher_col in range(3,max_col):
        current_teacher = Workbook()
        teacher_sheet=current_teacher.create_sheet('Data')
        new_col=1
        for i in range(4,(len((wb.get_sheet_names())))):
            ws = wb.worksheets[i]
            new_col=copy_columns(teacher_sheet,ws,new_col,teacher_col)
            
        teacher_name=ws.cell(row=1,column=teacher_col).value    
        current_teacher.save(teacher_name+'_.xlsx')
    return current_teacher


def copy_columns(teacher_sheet,ws,new_col,teacher_col):
    check=False
    num=0
    sub_count=0
    for row in range(2,ws.max_row):
        try:
            num=int(ws.cell(row=row,column=teacher_col).value)
        except:
            pass
        if check == False:
            sub_count=sub_count+1
            if num > 0:
                check=True
                teacher_sheet.cell(row=2,column=new_col,value=ws.cell(row=1,column=1).value)
                teacher_sheet.cell(row=2,column=new_col+1,value=ws.cell(row=1,column=2).value)
                teacher_sheet.cell(row=2,column=new_col+2,value=ws.cell(row=1,column=teacher_col).value)
        if check ==True:
            teacher_sheet.cell(row=row+2-sub_count,column=new_col,value=ws.cell(row=row,column=1).value)
            teacher_sheet.cell(row=row+2-sub_count,column=new_col+1,value=ws.cell(row=row,column=2).value)
            teacher_sheet.cell(row=row+2-sub_count,column=new_col+2,value=ws.cell(row=row,column=teacher_col).value)
        
    if check==True:
        return new_col+4
    else:
        return new_col
    
