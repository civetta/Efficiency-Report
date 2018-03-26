import openpyxl
from openpyxl import Workbook
from datetime import date
import calendar
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from copy import copy
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter


def create_books(max_col,wb):
    for teacher_col in range(3,max_col):
        current_teacher = Workbook()
        info=current_teacher.create_sheet('FAQ')
        teacher_sheet=current_teacher.create_sheet('Data')
        summary=current_teacher.create_sheet('Weekly Summary')
        copy_summary(summary,max_col,teacher_col,wb)
        img = Image('faq.png')
        info.add_image(img,'A1' )
        std=current_teacher.get_sheet_by_name('Sheet')
        current_teacher.remove_sheet(std)
        new_col=1
        for i in range(4,(len((wb.get_sheet_names())))):
            ws = wb.worksheets[i]
            new_col=copy_columns(teacher_sheet,ws,new_col,teacher_col)
            
        teacher_name=ws.cell(row=1,column=teacher_col).value    
        current_teacher.save(teacher_name+'_.xlsx')
    return current_teacher


def copy_summary(summary,max_col,teacher_col,wb):
    for i in range(4,(len((wb.get_sheet_names())))):
            ws = wb.worksheets[i]
    print ws.cell(row=1,column=teacher_col).value


def copy_columns(teacher_sheet,ws,new_col,teacher_col):
    check=False
    num=0
    sub_count=0
    grayFill=PatternFill("solid", fgColor="efefef")
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
                teacher_sheet.column_dimensions[get_column_letter(new_col)].width  =  int(20)
                teacher_sheet.cell(row=2,column=new_col+1,value=ws.cell(row=1,column=2).value)
                teacher_sheet.column_dimensions[get_column_letter(new_col+1)].width  =  int(20)
                teacher_sheet.cell(row=2,column=new_col+2,value=ws.cell(row=1,column=teacher_col).value)
                teacher_sheet.column_dimensions[get_column_letter(new_col+2)].width  =  int(20)
                teacher_sheet.cell(row=2,column=new_col+3).fill=grayFill
        if check ==True:
            teacher_sheet.cell(row=row+2-sub_count,column=new_col,value=ws.cell(row=row,column=1).value)
            teacher_sheet.cell(row=row+2-sub_count,column=new_col+1,value=ws.cell(row=row,column=2).value)
            new_cell=teacher_sheet.cell(row=row+2-sub_count,column=new_col+2)
            old_cell=ws.cell(row=row,column=teacher_col)
            new_cell.value=old_cell.value
            new_cell.font = copy(old_cell.font)
            new_cell.fill = copy(old_cell.fill)
            teacher_sheet.cell(row=row+2-sub_count,column=new_col+3).fill=grayFill
    if check==True:
        return new_col+4
    else:
        return new_col


def conditional_format(new_cell,old_cell):
    new
