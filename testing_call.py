from openpyxl import load_workbook
from daysplit2 import split_sheet_by_days
from blocker import define_blocks
from time_difference import make_time_difference_sheet
from create_tables import call_create_tables

"""User Input Variables"""
skip_days = ['03/29', '3/27']

"""Calling Functions"""
wb = load_workbook(filename='TestSource.xlsx')
make_time_difference_sheet(wb)
split_sheet_by_days(wb, skip_days)
tables_start_column = call_create_tables(wb)
define_blocks(wb,tables_start_column)
wb.save('Test2.xlsx')