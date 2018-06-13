from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
"""The abrivation actv is short for "active". It's used to talk about the 
current column, cell, or row in the summary page"""

def create_summary_page(wb,data_dict, checks):
    """Create the Summary page and locates it at the beginning of the workbook.
    Uses the sheet names to make date columns. Then it calls the first sheet
    and uses that to create the names."""
    night_color = 'c0b8f2'
    day_color = 'f2c0b8'
    sheet_list = wb.get_sheet_names()[:-2]
    num_of_days = len(sheet_list)
    wb.create_sheet('Summary', 0)
    summary = wb.get_sheet_by_name('Summary')
    rawsheet = wb.get_sheet_by_name('Raw Pulls')
    create_title(summary, sheet_list)
    if checks['Day Check'] is True and checks['Night Check'] is True:
        create_table(summary, rawsheet, 3, 'Day Summary', data_dict, num_of_days, day_color, wb)
        create_table(summary, rawsheet, 8+num_of_days, 'Night Summary', data_dict, num_of_days, night_color, wb)
    elif checks['Day Check'] is True and checks['Night Check'] is False:
        create_table(summary, rawsheet, 3, 'Day Summary', data_dict, num_of_days, day_color, wb)
    else:
        create_table(summary, rawsheet, 3, 'Night Summary', data_dict, num_of_days, night_color, wb)


def create_title(summary, sheet_list):
    """Creates and formats the title of the summary page. Uses 
    worksheet dates to create title"""
    date_range = sheet_list[0]+" - "+sheet_list[-1]
    title_of_summary = "Summary Of Days: " + date_range
    summary.cell(row=1, column=1).value = title_of_summary
    summary.cell(row=1, column=1).font = Font(size=30, bold=True)
    summary.row_dimensions[1].height = int(60)

def create_table(summary, rawsheet, header_row, table_name, data_dict, num_of_days, color, wb):
    """Uses the names from the first non summary page and copies and pastes
    them into the summary page, using a bit of formatting"""
    create_sub_titles(summary, header_row, table_name, num_of_days, color)
    fill_in_date_column(wb, summary, color, header_row, num_of_days) 
    for col_in_rawsheet in range(3, rawsheet.max_column):
        column_array = []
        current_col = col_in_rawsheet-1
        summary.column_dimensions[get_column_letter(current_col)].width = int(20)
        create_teacher_header_row(
            col_in_rawsheet, rawsheet, summary, header_row, current_col)
        for index, row in enumerate(range(header_row, header_row+num_of_days+1)):
            curr_cell = summary.cell(row=row, column=current_col)
            format_curr_cell(curr_cell, color)
            if index > 1:
                data = paste_data(
                    curr_cell, summary, current_col, row, header_row, data_dict, table_name)
                if data != '':
                    column_array.append(data)
        if len(column_array)>0:
            average = round(sum(column_array)/len(column_array),2)
            avg_cell_row = header_row+num_of_days+1
            average_cell = summary.cell(row=avg_cell_row, column=current_col)
            big_font(average_cell, average)
    summary.column_dimensions['A'].width = int(20)


def paste_data(curr_cell, summary, column, row, header_row, data_dict, table_name):
    date = summary.cell(row=row, column=1).value
    teacher_name = summary.cell(row=header_row, column=column).value
    teacher_name = teacher_name.replace('\r\n', ' ')
    date_dictionary = data_dict[date]
    teacher_data = date_dictionary[teacher_name]
    if "Day" in table_name:
        data = teacher_data['Day Average']
        curr_cell.value = data
    else:
        data = teacher_data['Night Average']
        curr_cell.value = data
    return data
    
    

def create_sub_titles(summary, header_row, table_name, num_of_days, color):
    summary.row_dimensions[header_row].height = int(30)
    title_of_table = summary.cell(row=header_row, column=1)
    format_curr_cell(title_of_table, color)
    total_average_title = summary.cell(row=header_row+num_of_days+1, column=1)
    big_font(title_of_table, table_name)
    big_font(total_average_title, 'Total Average')
    summary.row_dimensions[header_row].height = int(40)


def create_teacher_header_row(col_in_rawsheet, rawsheet, summary, header_row, current_col):
    teacher_name = rawsheet.cell(row=1, column=col_in_rawsheet).value
    first_name = teacher_name[:teacher_name.index(" ")]
    last_name = teacher_name[teacher_name.index(" ")+1:]
    teacher_name_formatted = first_name+'\r\n'+last_name
    current_cell = summary.cell(row=header_row, column=current_col)
    big_font(current_cell, teacher_name_formatted)


def fill_in_date_column(wb, summary, color, header_row, num_of_days):
    dates = wb.get_sheet_names()[1:-2]
    count = 0
    for row in range(header_row+1, header_row+num_of_days+1):
        date_cell = summary.cell(row=row, column=1)
        date_cell.value = dates[count]
        format_curr_cell(date_cell, color)
        count = count+1 


def format_curr_cell(curr_cell, color):
    thin_border = Border(left=Side(border_style='thin', color='E6E6E6'),
        right=Side(border_style='thin', color='E6E6E6'),
        top=Side(border_style='thin', color='E6E6E6'),
        bottom=Side(border_style='thin', color='E6E6E6'))
    curr_cell.fill = PatternFill("solid", fgColor=color)
    curr_cell.border = thin_border


def big_font(cell, value):
    cell.value = value
    cell.font = Font(size=15, bold=True)
    cell.alignment = Alignment(wrapText=True)
