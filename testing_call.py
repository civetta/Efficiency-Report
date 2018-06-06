from openpyxl import load_workbook
from time_difference import make_time_difference_sheet
from daysplit2 import split_sheet_by_days
from blocker import define_blocks
"""User Input Variables"""
skip_days = ['03/29', '3/27']

"""Calling Functions"""
wb = load_workbook(filename='TestSource.xlsx')
make_time_difference_sheet(wb)
split_sheet_by_days(wb, skip_days)
define_blocks(wb)
wb.save('Test2.xlsx')


