from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

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
    #color_fill(3,len(sheet_list))

#def color_fill(row_to_start,)
def create_title(summary, sheet_list):
    """Creates and formats the title of the summary page. Uses 
    worksheet dates to create title"""
    date_range = sheet_list[0]+" - "+sheet_list[-1]
    title_of_summary = "Summary Of Days: " + date_range
    summary.cell(row=1, column=1).value = title_of_summary
    summary.cell(row=1, column=1).font = Font(size=20, bold=True)
    summary.column_dimensions['A'].width = int(40)
    summary.row_dimensions[1].height = int(60)


def create_table(summary, sheet_1, row_to_paste, name_of_table,data_dict):
    """Uses the names from the first non summary page and copies and pastes
    them into the summary page, using a bit of formatting"""
    summary.row_dimensions[row_to_paste].height = int(30)
    summary.cell(row=row_to_paste, column=1, value=name_of_table)
    summary.cell(row=row_to_paste, column=1).font = Font(size=15, bold=True)
    summary.cell(row=row_to_paste, column=1).alignment = Alignment(wrapText=True)
    for column_in_sheet1 in range(3, sheet_1.max_column):
        column_in_sum = column_in_sheet1-1
        teacher_name = sheet_1.cell(row=1, column=column_in_sheet1).value
        first_name = teacher_name[:teacher_name.index(" ")]
        last_name = teacher_name[teacher_name.index(" ")+1:]
        teacher_name_formatted = first_name+'\r\n'+last_name
        summary.cell(row=row_to_paste, column=column_in_sum, value = teacher_name_formatted)
        summary.cell(row=row_to_paste, column=column_in_sum).alignment = Alignment(wrapText=True)
        summary.column_dimensions[get_column_letter(column_in_sum)].width = int(15)
    



def check_which_tables_to_make(wb, data_dict, summary, sheet1, number_of_days):
    day_check = False
    night_check = False
    wb
    for day in data_dict.keys():
        day_dict = data_dict[day]
        for teacher in day_dict.keys():
            current_dict = day_dict[teacher]
            if current_dict['Day Average'] is not None:
                day_check = True
            if current_dict['Night Average'] is not None:
                night_check = True
    if day_check == True and night_check == True:
        create_table(summary, sheet1, 3, 'Day Summary',data_dict)
        create_table(summary, sheet1, 6+number_of_days, 'Night Summary',data_dict)
    elif day_check == True and night_check == False:
        create_table(summary, sheet1, 3, 'Day Summary',data_dict)
    else:
        create_table(summary, sheet1, 3, 'Night Summary',data_dict)