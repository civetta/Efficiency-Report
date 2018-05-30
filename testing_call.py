import openpyxl
from openpyxl import load_workbook
import numpy as np
import re
import os
from daysplit2 import split_sheet_by_days
from blocker import define_blocks
from efficency_calculator import create_block_table
from break_down import create_summary
from weekly_summary import find_row
from time_difference import make_time_difference_sheet
from Archive import archive_to_excel
from formatter import formatter
from teacher_books import create_books



wb = load_workbook(filename = 'TestSource.xlsx')
print "Start Time Difference"
wb=make_time_difference_sheet(wb)
print "Difference Done"
wb=split_sheet_by_days(wb,'03/29')
print "Split Done"
wb=define_blocks(wb)
print "Blocking Done"
wb.save('Test2.xlsx')