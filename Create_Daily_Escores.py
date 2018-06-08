from Past_Data_5B import find_empty_row
from openpyxl.styles import Font


def find_non_empty_tables(wb):
    week = wb.get_sheet_names()
    week = week[:-3]
    for day in week:
        ws = wb.get_sheet_by_name(day)
        for col in range(8, ws.max_column+1):
            table_row = ((col-7)*8)-6
            empty_row = find_empty_row(ws, table_row)
            if table_row-empty_row != 0:
                create_arrays(ws, table_row, empty_row)


def create_arrays(ws, table_row, empty_row):
    day_array = []
    night_array = []
    for row in range(table_row+1, empty_row):
        block_escore = ws.cell(row=row, column=4).value
        if '*' in ws.cell(row=row, column=1).value:
            night_array.append(block_escore)
        else:
            day_array.append(block_escore)
    past_daily_escore(ws, empty_row, day_array, night_array)


def past_daily_escore(ws, empty_row, day_array, night_array):
    if len(day_array) > 0:
        day_average = sum(day_array)/len(day_array)
        ws.cell(row=empty_row, column=1, value="Day Average").font = Font(bold=True)
        ws.cell(row=empty_row, column=4, value=day_average).font = Font(bold=True)
        empty_row = empty_row+1
    if len(night_array) > 0:
        night_average = sum(night_array)/len(night_array)
        ws.cell(row=empty_row, column=1, value="Night Average").font = Font(bold=True)
        ws.cell(row=empty_row, column=4, value=night_average).font = Font(bold=True)
