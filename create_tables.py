from openpyxl.styles import PatternFill
import openpyxl

def create_summary_tables(ws, max_col):
    """Creates Daily Summary Tables for each teacher to the right of the data""" 
    for col in range(3, max_col+1):
        teacher_name = ws.cell(row=1, column=col).value
        ws.cell(row=col*8-20, column=max_col+2, value=teacher_name)
        ws.cell(row=col*8-20+1, column=max_col+2, value="Time")
        ws.cell(row=col*8-20+1, column=max_col+3, value="Average Students")
        ws.cell(row=col*8-20+1, column=max_col+4, value="Average Tabby")
        ws.cell(row=col*8-20+1, column=max_col+5, value="Effciency Score")
    
def create_team_daily_table(ws, max_col, shift, start_row, hex_code):
    """Creates the Team Wide Summary Table. Is called twice.
    Once to create a "Day Time" table and another to create
    the "night time" table"""  
    DayName = ws.cell(row=5, column=1).value[9:12]
    title = ws.cell(row=1+start_row, column=max_col+8)
    name = ws.cell(row=2+start_row, column=max_col+8)

    average_title = ws.cell(row=3+start_row, column=max_col+8)
    title.value = DayName+" "+shift+" Summary"
    name.value = 'Teacher Name'
    average_title.value = DayName+" Daily Average"

    title.fill = PatternFill("solid", fgColor=hex_code)
    name.fill = PatternFill("solid", fgColor=hex_code)
    average_title.fill = PatternFill("solid", fgColor=hex_code)
    for col in range(3, max_col+1):
        teach_name = ws.cell(row=1, column=col).value
        cell = ws.cell(row=2+start_row, column=max_col+col+6, value=teach_name)
        cell.fill = PatternFill("solid", fgColor=hex_code)