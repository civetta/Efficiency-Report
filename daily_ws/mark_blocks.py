from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from calculate_block_escore import organize_data


def define_blocks(wb, checks, scores):
    """Goes through each ws and then goes through each column looking for the 
    start and end of each block. It then calls the bolder function
    conditionally format the cells, and then passes information over to 
    calculate_block_escore module to do calculations and paste data into 
    tables"""
    week = wb.get_sheet_names()
    week = week[:-1]
    for day in week:
        ws = wb.get_sheet_by_name(day)
        max_col = ws.max_column
        max_row = find_max_row(ws)
        col = 8
        start_row_to_look = 2
        while col <= max_col:
            start = find_blocks(ws, col, max_row, start_row_to_look, 'start')
            if start != 'Next_Col' and start >= start_row_to_look and start != max_row :
                end = find_blocks(ws, col, max_row, start, 'end')
                start_row_to_look = end
                safe_to_color = empty_tabby(start,end,ws)
                if safe_to_color is True:
                    checks = bolder(ws, start, end, col, max_col, checks, scores,wb)                  
            else:
                col = col+1
                start_row_to_look = 2
    return checks


def find_max_row(ws):
    for row in range(1,ws.max_row):
        row_value = ws.cell(row=row,column=7).value
        if row_value == None:
            return row

            
def empty_tabby(start,end,ws):
    tab_list = []
    for r in range(start, end):
        Tabby_Cell = ws.cell(row=r,  column=7).value
        if Tabby_Cell != None:
            Tabby_Cell = int(Tabby_Cell)
        tab_list.append(Tabby_Cell)
    average_tabby = round(sum(tab_list)/float(len(tab_list)), 2)
    if average_tabby>0:
        return True
    else:
        print tab_list
        return False

def find_blocks(ws, col, max_row, starting_row, position):
    """Looks for either two 0's in a row for the end of a block,or 
    two sequential non zeros for the start of the block"""
    for row in range(starting_row, max_row):
        val1 = ws.cell(row=row, column=col).value
        if val1==None:
            return 'Next_Col'
        if position == 'start':  
            if int(val1) > 0:
                val2 = int(ws.cell(row=row+1, column=col).value)
                val3 = int(ws.cell(row=row+2, column=col).value)
                if val2 > 0 or val3 > 0:
                    return row
        elif position == 'end':
            if int(val1) == 0:
                try:
                    val2 = int(ws.cell(row=row+1, column=col).value)
                    val3 = int(ws.cell(row=row+2, column=col).value)
                except TypeError:
                    break
                if val2 == 0 and val3 == 0:
                    return row
    return max_row


def bolder(ws, start, end, column, max_col, checks, scores,wb):
    """This function goes through and bolds and conditionally
    formatts each of the blocks. It then creates a list of 
    all of the information in the blocks. So Tab_list is a 
    list of all of the tabby's during the block, block_list is a list
    of all of the students taken during that block, and so on. This
    information is passed over into the calculate_block_escore module
    who pastes all of these data into daily teacher summary tables."""            
    tab_list = []
    block_list = []
    for r in range(start, end):
        current_cell = ws.cell(row=r,  column=column)
        current_value = current_cell.value
        Tabby_Cell = int(ws.cell(row=r,  column=7).value)
        plus_1_check = current_value == Tabby_Cell+1 
        minus_1_check = current_value == Tabby_Cell-1
        current_cell.font = Font(bold=True)
        if current_value == Tabby_Cell or plus_1_check or minus_1_check:
            current_cell.fill = PatternFill("solid", fgColor='dff7c0')
        elif current_value < Tabby_Cell and current_value > 0:
            current_cell.fill = PatternFill("solid", fgColor='f2b8ea')
        elif current_value >= Tabby_Cell+2:
            current_cell.fill = PatternFill("solid", fgColor='c0f7f4')
        else:
                continue
        tab_list.append(Tabby_Cell)
        block_list.append(current_value)
    average_tabby = round(sum(tab_list)/float(len(tab_list)), 2)
    checks = organize_data(
        ws, start, end, column, block_list, tab_list, max_col, checks, scores,wb)
    return checks
