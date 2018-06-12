from Time_Difference_03 import make_time_difference_sheet
from Mark_Blocks_05 import define_blocks
from Day_Split_04 import split_sheet_by_days
from openpyxl import load_workbook
from Create_Daily_Escores import find_non_empty_tables
from efficiency_score_summary import create_summary_page
"""User Input Variables"""
skip_days = ['03/29', '3/27']
condition_list = {"Good Score": float(.90), "Upper Bound": float(1.25)}

"""Calling Functions"""
wb = load_workbook(filename='TestSource.xlsx')
make_time_difference_sheet(wb)
wb.save('01.xlsx')
split_sheet_by_days(wb, skip_days)
wb.save('02.xlsx')
lister = wb.get_sheet_names
print lister
define_blocks(wb)
wb.save('03.xlsx')
data_library = find_non_empty_tables(wb)
wb.save('04.xlsx')
#blank_sheeet = wb.get_sheet_by_name('Sheet')
#wb.remove_sheet(blank_sheeet)
#create_summary_page(wb, data_library)
wb.save('Test3.xlsx')