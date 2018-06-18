
from daily_ws.format_block_escore import find_table, find_empty_row
from utility import copier, find_teacher_column
from summary_ws.create_summary_tables import big_font

def copy_summary(teacherbook, wb, teachername):
    """Goes through each day of the week worksheet and using the 
    "find_table function that is used earlier when calculating daily
    blocks. The find table function returns the row in which teachername
    daily summary table starts. Then it copies the first five columns and 8 rows
    since when creating the tables, they were all 8 rows tall."""
    teach_summary = teacherbook.create_sheet('Weekly Summary')
    copy_daily_tables(teach_summary, wb, teachername)
    copy_summary_page(teach_summary, wb, teachername)
    format_summary_page(teach_summary)


def copy_daily_tables(teach_summary, wb, teachername):
    days_in_sheet = wb.get_sheet_names()[1:-2]
    teacherbook_row = 5
    for day in days_in_sheet:
        ws = wb.get_sheet_by_name(day)
        max_col = ws.max_column
        start_of_table = find_table(ws, teachername, max_col)
        date = ws.cell(row=2, column=6).value
        date = date[:date.index(" ")]
        date_cell = teach_summary.cell(row=teacherbook_row-1, column=1) 
        big_font(date_cell, date)
        for row in range(start_of_table, start_of_table+8):
            for col in range(1, 5):
                old_cell = ws.cell(row=row, column=col)
                new_cell = teach_summary.cell(row=teacherbook_row, column=col)
                copier(old_cell, new_cell)
            teacherbook_row = teacherbook_row+1


def copy_summary_page(teach_summary, wb, teachername):
    old_summary = wb.get_sheet_by_name('Summary')
    teacher_col = find_teacher_column(old_summary, teachername, 3, 1)
    title = old_summary.cell(row=1, column=1)
    new_title = teach_summary.cell(row=1, column=1)
    copier(title, new_title)
    day_table_end = find_empty_row(old_summary, 3)
    copy_sum_tables(
        3, day_table_end, teacher_col, old_summary, teach_summary, 6)
    if day_table_end < old_summary.max_row:
        night_table_end = find_empty_row(old_summary, day_table_end+2)
        copy_sum_tables(
            day_table_end+2, night_table_end, teacher_col, old_summary, teach_summary, 9)


def copy_sum_tables(start,end, teacher_col, old_summary, teach_summary, start_column):
    for row in range(start, end+1):
        old_date = old_summary.cell(row=row, column=1)
        new_date = teach_summary.cell(row=row, column=start_column)
        copier(old_date, new_date)
        old_cell = old_summary.cell(row=row, column=teacher_col)
        new_cell = teach_summary.cell(row=row, column=start_column+1)
        copier(old_cell, new_cell)
        


def format_summary_page(teach_summary):
    teach_summary.column_dimensions['A'].width = int(30)
    teach_summary.column_dimensions['B'].width = int(20)
    teach_summary.column_dimensions['C'].width = int(20)
    teach_summary.column_dimensions['D'].width = int(20)
    teach_summary.column_dimensions['F'].width = int(30)
    teach_summary.column_dimensions['G'].width = int(30)
    

        

