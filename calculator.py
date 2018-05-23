import openpyxl
import numpy as np


"""Declares all of the variables needed to create and organize the efficiency table"""
def organize_data(ws,start,end,column,block,tab_list,max_col):
    average_student= round(sum(block)/float(len(block)),2)
    average_tabby=round(sum(tab_list)/float(len(tab_list)),2)
    block_escore=round(average_student/float(average_tabby),2)
    teacher_name=str(ws.cell(row=1,column=column).value)
    start_time=str(ws.cell(row=start,column=1).value)
    end_time=str(ws.cell(row=end,column=1).value)
    time_range=create_time_range(start_time,end_time)
    paste_list=[time_range,average_student,average_tabby,block_escore]
    start_col=max_col+2
    paste_data(ws,paste_list,teacher_name,column,start_col)

"""Creates a time range from the date cell. So it reads 2PM-4PM instead of Tuesday, May 21, 2PM-Tuesaday May 21, 4PM"""
def create_time_range(start,end):
    lister=[start,end]
    time_range=""
    for item in lister:
        item=item[item.index(" ")+1:]
        item=item[item.index(" ")+1:]
        time_range=time_range+item+" - "
    return time_range[:-3]

"""Paste the organized data under the daily summary tables"""        
def paste_data(ws,paste_list,teacher_name,column,start_col):
    start_row=find_start_row(ws,teacher_name)
    for count in range(len(paste_list)):
        ws.cell(row = start_row+1,column=start_col+count).value=paste_list[count]


def find_start_row(ws,teacher_name):
    for row in range(1,ws.max_row):
        if ws.cell(row=row,column=start_col).value ==teacher_name:
            start_row=row
            return start_row
        
def create_summary_tables(ws,max_col):
    for col in range(2,max_col):
        teacher_name=ws.cell(row=1,column=col).value
        ws.cell(row=col*7-7,column=max_col+2).value=teacher_name
        ws.cell(row=col*7-7+1,column=max_col+2).value="Time"
        ws.cell(row=col*7-7+1,column=max_col+3).value="Average Students"
        ws.cell(row=col*7-7+1,column=max_col+4).value="Average Tabby"
        ws.cell(row=col*7-7+1,column=max_col+5).value="Effciency Score"
    
    
    
    
    
