import openpyxl
import numpy as np
from datetime import datetime
from openpyxl.styles import Fill
from openpyxl.styles import PatternFill


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

"""
Input: Two string of time. Example: 03/21/18 Thu 8:27 AM and 03/21/18 Thu 9:48 AM.
Output: A string with just the hours (example: 8:27 AM - 9:48 AM").
Times is marked with an asteriks if one of the times is after 8 AM but before 1AM.
Or if the day of the week is Saturday or Sunday
These mark the cells so that they are calculated differently then the "day time" teachers.
"""
def create_time_range(start,end):
    lister=[start,end]
    time_range=""
    night_shift_indicator_start = datetime.strptime('8:00PM', "%I:%M%p")
    night_shift_indicator_end = datetime.strptime('2:00AM', "%I:%M%p")
    for item in lister:
        hour=item[13:]
        date=item[9:12]
        string_time=datetime.strptime(hour,'%I:%M %p')
        if date =='Sat' or date=='Sun' or night_shift_indicator_start <= string_time or string_time<=night_shift_indicator_end:
            hour=hour+'*'
        time_range=time_range+hour+" - "
    return time_range[:-3]

"""Paste the organized data under the daily summary tables.
It also makes the cells purple if it has an asteriks(*) which indiciates a night or weekend shift."""        
def paste_data(ws,paste_list,teacher_name,column,start_col):
    if teacher_name != "":
        table_row=find_table(ws,teacher_name,start_col)
        empty_row=find_empty_row(ws,table_row,start_col)
        if empty_row != None:
            if paste_list[0].find('*')>-1:
                paste_list[0].replace('*','')
                hex_color="c6c0ed"
            else:
                hex_color="ffffff"
            for count in range(len(paste_list)):
                ws.cell(row = empty_row,column=start_col+count,value=paste_list[count]).fill=PatternFill("solid", fgColor=hex_color)

"""Finds the location of the teacher's daily summary table in the column"""
def find_table(ws,teacher_name,start_col):
    for col in range(1,start_col):
        if ws.cell(row=1,column=col).value ==teacher_name:
            return col*8-20
        
"""Finds the first empty row in the teacher's daily summary table so it can paste the new data in it"""
def find_empty_row(ws,table_row,start_col):
    for row in range(table_row,ws.max_row):
        if ws.cell(row=row,column=start_col).value==None:
            return row
        
"""Creates Daily Summary Tables for each teacher to the right of the data"""        
def create_summary_tables(ws,max_col):
    for col in range(3,max_col+1):
        teacher_name=ws.cell(row=1,column=col).value
        ws.cell(row=col*8-20,column=max_col+2).value=teacher_name
        ws.cell(row=col*8-20+1,column=max_col+2).value="Time"
        ws.cell(row=col*8-20+1,column=max_col+3).value="Average Students"
        ws.cell(row=col*8-20+1,column=max_col+4).value="Average Tabby"
        ws.cell(row=col*8-20+1,column=max_col+5).value="Effciency Score"
    
"""Creates the Team Wide Summary Table. Is called twice.
Once to create a "Day Time" table and another to create the "night time" table"""  
def create_team_daily_table(ws,max_col,shift,start_row,hex_code):
    DayName=ws.cell(row=5,column=1).value[9:12]
    ws.cell(row=1+start_row,column=max_col+8,value=DayName+" "+shift+" Summary").fill=PatternFill("solid", fgColor=hex_code)
    ws.cell(row=2+start_row,column=max_col+8,value="Teacher Name").fill=PatternFill("solid", fgColor=hex_code)
    ws.cell(row=3+start_row,column=max_col+8,value=DayName+" Daily Average").fill=PatternFill("solid", fgColor=hex_code)
    for col in range(3,max_col+1):
        teacher_name=ws.cell(row=1,column=col).value
        ws.cell(row=2+start_row,column=max_col+col+6,value=teacher_name).fill=PatternFill("solid", fgColor=hex_code)
    

    

