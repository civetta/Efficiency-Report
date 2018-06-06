from openpyxl.styles import PatternFill


def call_create_tables(wb):
    week = wb.get_sheet_names()
    week = week[3:]
    for day in week:
        ws = wb.get_sheet_by_name(day)
        max_col = ws.max_column
        create_summary_tables(ws, max_col)
    return max_col


def create_summary_tables(ws, max_col):
    """Creates Daily Summary Tables for each teacher
     to the right of the data""" 
    for col in range(8, max_col+1):
        teacher_name = ws.cell(row=1, column=col).value
        ws.cell(row=col*8-20, column=1, value=teacher_name)
        ws.cell(row=col*8-20+1, column=1, value="Time")
        ws.cell(row=col*8-20+1, column=2, value="Average Students")
        ws.cell(row=col*8-20+1, column=4, value="Average Tabby")
        ws.cell(row=col*8-20+1, column=max_col+5, value="Effciency Score")