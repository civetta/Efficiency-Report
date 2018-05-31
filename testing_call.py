from openpyxl import load_workbook
from daysplit2 import split_sheet_by_days
from blocker import define_blocks
from time_difference import make_time_difference_sheet
from create_tables import call_create_tables

wb = load_workbook(filename='TestSource.xlsx')
wb = make_time_difference_sheet(wb)
wb = split_sheet_by_days(wb, ['03/29','3/27'])
wb.save('Test2.xlsx')
#wb = call_create_tables(wb)
#wb = define_blocks(wb)
