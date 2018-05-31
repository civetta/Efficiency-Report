from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from calculator import organize_data
from calculator import daily_array
from create_tables import create_summary_tables
from create_tables import create_team_daily_table
import re


def create_list(ws, col, max_row):
    """Starting in row 2, it creates an array for the column"""
    column_list = []

    for i in range(1, max_row+1):
        valv = ws.cell(row=i, column=col).value
        try:
            valv = int(valv)
            column_list.append(valv)
        except:
            continue
    return column_list


def find_teacher_chunks(lister):
    """Finds the numbers of chunks and about where they are, 
    then uses trim, to define them exactly"""
    o_lister = str(lister)
    lister = o_lister.replace('0, 0','-')
    list_of_blocks = re.split('-,', lister)
    final_array = []
    for i in list_of_blocks:
        try:
            a = trim(i)
            final_array.append(a)
        except:
            continue
    return final_array


def trim(lister):
    """Removes trailing and begining zeros from the chunks"""
    start = re.search('[1-9]', lister).start()
    lister = lister[start:]
    lil_list = re.findall('[1-9]', lister)
    lister = lister[:lister.rfind(lil_list[-1])+1]
    return lister


def chunk_location(lister, o_lister, ws, col, max_col):
    """Identifies the start of the chunk location in the orginal
     array which is the entire column +2. So the third item in o_lister 
     is actually in row 4"""
    for sub_list in lister:
        sub_list = integer_list(sub_list)
        for i in range(len(o_lister)):
            same_length_boolean = o_lister[i:i+len(sub_list)] == sub_list
            if same_length_boolean and len(sub_list)>2 and sum(sub_list)>2:
                bolder(ws, (i+2), (i+2+len(sub_list)), col, sub_list, max_col)


def integer_list(lister):
    """This function takes the list of strings"""        
    new_lister = []
    lister = lister.replace(" ", "").split(",")
    for i in lister:
        new_lister.append(int(i))
    return new_lister


def bolder(ws, start_index, end_index, column, block, max_col):
    """This function goes through and bolds and conditionally 
    formatts each of the blocks. The bolding will be used later to identify 
    a block and for calculating"""            
    tab_list = []
    
    for r in range(start_index, end_index):
        Tabby_Cell = ws.cell(row=r,  column=2).value
        Tabby_Cell = int(Tabby_Cell)
        tab_list.append(Tabby_Cell)
        current_cell = ws.cell(row=r,  column=column)
        current_value = current_cell.value
        plus_1_check = current_value == Tabby_Cell+1 
        minus_1_check = current_value == Tabby_Cell-1
        if current_value == 0:
            current_cell.fill = PatternFill("solid", fgColor='ffffff')
            current_cell.font = Font(bold=True)
        elif current_value == Tabby_Cell or plus_1_check or minus_1_check:
            current_cell.fill = PatternFill("solid", fgColor='dff7c0')
            current_cell.font = Font(bold=True)
        elif current_value < Tabby_Cell and current_value > 0:
            current_cell.fill = PatternFill("solid", fgColor='f2b8ea')
            current_cell.font = Font(bold=True)
        elif current_value >= Tabby_Cell+2:
            current_cell.fill = PatternFill("solid", fgColor='c0f7f4')
            current_cell.font = Font(bold=True)
        else:
                continue
    organize_data(ws, start_index, end_index, column, block, tab_list, max_col)


def define_blocks(wb):
    week = wb.get_sheet_names()
    week = week[3:]
    for day in week:
        ws = wb.get_sheet_by_name(day)
        max_row = ws.max_row
        max_col = ws.max_column
        counter = max_col
        i = 3
        while i <= counter:
            lister = create_list(ws, i, max_row)
            new_list = find_teacher_chunks(lister)
            chunk_location(new_list, lister, ws, i,max_col)
            i = i+1
    return wb


