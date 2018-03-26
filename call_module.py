import openpyxl
from openpyxl import load_workbook
import numpy as np
import re
import os
from day_split import find_days
from blocker import define_blocks
from efficency_calculator import create_block_table
from break_down import create_summary
from weekly_summary import find_row
from time_difference import make_sheet
from Archive import call_function
from formatter import formatter
from teacher_books import create_books

Lead_Names=["Caren Glowa",'']
#"Jeremy Shock","Rachel Adams","Jairo  Rios","Salome Saenz","Kristin Donnelly",
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
    folder_name = date+"--EReport"
    folder_location = os.path.join('C:\Users\kheyden\Documents\Program\2017\WeeklySummary', folder_name)
    #folder_location = os.path.join('C:\Users\kheyden\OneDrive - Imagine Learning\Efficiency Report', folder_name)
    if not os.path.exists(folder_location):
        os.makedirs(folder_location)
    final_save_name = os.path.join(folder_location,Lead_Names[x]+"_"+date+".xlsx")
    wb.save(final_save_name)
    wb.save("THEWORKBOOK.xlsx")
    
    teacher=create_books(max_col,wb)
    teacher.save(final_save_name)

    

