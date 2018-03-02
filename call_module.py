import openpyxl
from openpyxl import load_workbook
import numpy as np
import re
from day_split import find_days
from blocker import define_blocks
from efficency_calculator import create_block_table
from break_down import create_summary
from weekly_summary import find_row
from time_difference import make_sheet
from Archive import call_function
from formatter import formatter

Lead_Names=["Jeremy Shock","Rachel Adams","Jairo  Rios","Salome Saenz","Kristin Donnelly","Caren Glowa",'']
#
for x in range(len(Lead_Names)-1):
    print "------------------------------------------------------------"
    print Lead_Names[x]
    start=Lead_Names[x]
    end=Lead_Names[x+1]
    wb = call_function(start,end)
    print "Archived Sheet Completed"


    wb=make_sheet(wb)
    print "Difference Sheet Completed"


    wb=find_days(wb)
    
    print "Days of Week Sheets Completed"
    wb=define_blocks(wb)
    print "Blocks Bolded and Conditional Formatted"
    max_col=create_block_table(wb,start)
    print "Blocks Calculated"
    create_summary(wb,max_col)
    print "Summary Breakdown Completed"
    find_row(wb)
    print "Final Summary Page Completed"
    print "Program Completed"
    std=wb.get_sheet_by_name('Sheet')
    wb.remove_sheet(std)
    formatter(wb)
    raw_sheet = wb.get_sheet_by_name("Raw Changes")
    first_date=str(raw_sheet.cell(row=2,column=1).value)
    date=first_date[:first_date.index(" ")]
    date=date.replace("/","-")
    wb.save(Lead_Names[x]+"_"+date+".xlsx")
    print "\n"

    

