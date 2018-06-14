from openpyxl import load_workbook
from time_difference import make_time_difference_sheet
from split_days import split_sheet_by_days
from create_tables import call_create_tables
from mark_blocks import define_blocks
from calculate_daily_escore import find_non_empty_tables
from efficiency_score_summary import create_summary_page
"""User Input Variables"""
skip_days = ['03/29', '3/27']
scores = {"Good Score": float(.90), "Upper Bound": float(1.25)}
output_filename = "Hey"


"""Calling Functions"""
wb = load_workbook(filename='FullTime_Team_Source.xlsx')
blank_sheeet = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(blank_sheeet)
make_time_difference_sheet(wb)
split_sheet_by_days(wb, skip_days)
call_create_tables(wb)
checks = {'Night Check': False, 'Day Check': False}
checks = define_blocks(wb, checks, scores)
data_library = find_non_empty_tables(wb)
create_summary_page(wb, data_library, checks)
wb.save(output_filename+'.xlsx')



