from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side


def create_summary_page(wb,data_dict):
    """Create the Summary page and locates it at the beginning of the workbook.
    Uses the sheet names to make date columns. Then it calls the first sheet
    and uses that to create the names."""
    sheet_list = wb.get_sheet_names()[:-2]
    wb.create_sheet('Summary', 0)
    summary = wb.get_sheet_by_name('Summary')
    sheet1 = wb.get_sheet_by_name('Raw Pulls')
    create_title(summary, sheet_list)
    check_which_tables_to_make(wb, data_dict, summary, sheet1,len(sheet_list))


def create_title(summary, sheet_list):
    """Creates and formats the title of the summary page. Uses 
    worksheet dates to create title"""
    date_range = sheet_list[0]+" - "+sheet_list[-1]
    title_of_summary = "Summary Of Days: " + date_range
    summary.cell(row=1, column=1).value = title_of_summary
    summary.cell(row=1, column=1).font = Font(size=20, bold=True)
    summary.column_dimensions['A'].width = int(40)
    summary.row_dimensions[1].height = int(60)


def check_which_tables_to_make(wb, data_dict, summary, sheet1, number_of_days):
    """Reads through the data dictionary and sees if there are night time and 
    day time shifts. If there are both it creates tables for both."""
    day_check = False
    night_check = False
    night_color = 'c0b8f2'
    day_color = 'f2c0b8'
    for day in data_dict.keys():
        day_dict = data_dict[day]
        for teacher in day_dict.keys():
            current_dict = day_dict[teacher]
            if current_dict['Day Average'] is not None:
                day_check = True
            if current_dict['Night Average'] is not None:
                night_check = True
    if day_check == True and night_check == True:
        create_table(summary, sheet1, 3, 
            'Day Summary',data_dict, number_of_days, day_color, wb)
        create_table(summary, sheet1, 8+number_of_days, 
            'Night Summary', data_dict, number_of_days, night_color, wb)
    elif day_check == True and night_check == False:
        create_table(summary, sheet1, 3, 
            'Day Summary',data_dict, number_of_days, day_color, wb)
    else:
        create_table(summary, sheet1, 3, 
            'Night Summary',data_dict, number_of_days, night_color, wb)


def create_table(summary, sheet_1, header_row, name_of_table, data_dict, number_of_days, color, wb):
    """Uses the names from the first non summary page and copies and pastes
    them into the summary page, using a bit of formatting"""
    create_titles(summary, header_row, name_of_table, number_of_days, color)
    fill_in_date_column(wb, summary, color, header_row, number_of_days) 
    for column_in_sheet1 in range(3, sheet_1.max_column):
        column_in_sum = column_in_sheet1-1
        create_teacher_name_column_head(column_in_sheet1, sheet_1, summary, header_row, column_in_sum)
        for index, row in enumerate(range(header_row, header_row+number_of_days+1)):
            cell_in_table = summary.cell(row=row, column=column_in_sum)
            format_cell_in_table(cell_in_table, color)
            if index > 1:
                paste_data_into_cell(cell_in_table, summary, column_in_sum, row, header_row, data_dict, name_of_table)


def paste_data_into_cell(cell_in_table, summary, column, row, header_row, data_dict, name_of_table):
    date = summary.cell(row=row, column=1).value
    teacher_name = summary.cell(row=header_row, column=column).value
    teacher_name = teacher_name.replace('\r\n', ' ')
    date_dictionary = data_dict[date]
    teacher_data = date_dictionary[teacher_name]
    if "Day" in name_of_table:
        cell_in_table.value = teacher_data['Day Average']
    else:
        cell_in_table.value = teacher_data['Night Average']
    
    
    

def create_titles(summary, header_row, name_of_table, number_of_days, color):
    summary.row_dimensions[header_row].height = int(30)
    title_of_table = summary.cell(row=header_row, column=1)
    format_cell_in_table(title_of_table, color)
    total_average_title = summary.cell(row=header_row+number_of_days+1, column=1)
    format_table_title_text(title_of_table, name_of_table)
    format_table_title_text(total_average_title, 'Total Average')


def create_teacher_name_column_head(column_in_sheet1, sheet_1, summary, header_row, column_in_sum):
    teacher_name_formatted = create_formatted_teacher_name(sheet_1, column_in_sheet1)
    current_cell = summary.cell(row=header_row, column=column_in_sum)
    paste_formated_teacher_name(teacher_name_formatted, column_in_sum, current_cell, summary)


def fill_in_date_column(wb, summary, color, header_row, number_of_days):
    dates = wb.get_sheet_names()[1:-2]
    count = 0
    for row in range(header_row+1, header_row+number_of_days+1):
        date_cell = summary.cell(row=row, column=1)
        date_cell.value = dates[count]
        format_cell_in_table(date_cell, color)
        count = count+1 


def format_cell_in_table(cell_in_table, color):
    thin_border = Border(left=Side(border_style='thin', color='E6E6E6'),
        right=Side(border_style='thin', color='E6E6E6'),
        top=Side(border_style='thin', color='E6E6E6'),
        bottom=Side(border_style='thin', color='E6E6E6'))
    cell_in_table.fill = PatternFill("solid", fgColor=color)
    cell_in_table.border = thin_border


def create_formatted_teacher_name(sheet_1, column_in_sheet1):
    teacher_name = sheet_1.cell(row=1, column=column_in_sheet1).value
    first_name = teacher_name[:teacher_name.index(" ")]
    last_name = teacher_name[teacher_name.index(" ")+1:]
    teacher_name_formatted = first_name+'\r\n'+last_name
    return teacher_name_formatted


def paste_formated_teacher_name(teacher_name_formatted, column_in_sum, current_cell, summary):
    current_cell.value = teacher_name_formatted
    current_cell.alignment = Alignment(wrapText=True)
    current_cell.font = Font(size=10, bold=True)
    summary.column_dimensions[get_column_letter(column_in_sum)].width = int(15)


def format_table_title_text(cell, value):
    cell.value = value
    cell.font = Font(size=15, bold=True)
    cell.alignment = Alignment(wrapText=True)