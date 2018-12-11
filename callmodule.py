from openpyxl import load_workbook
from periscope_source import create_input
from time_difference import make_time_difference_sheet
from split_days import split_sheet_by_days
from daily_ws.create_tables import call_create_tables
from daily_ws.mark_blocks import define_blocks
from summary_ws.calculate_daily_escore import find_non_empty_tables
from summary_ws.efficiency_score_summary import create_summary_page
from teacherbooks.create_teacher_books import create_books
from summary_ws.power_bi_format import create_dataframe
from datetime import datetime
import os
import warnings
warnings.filterwarnings("ignore")


def save_leadbook(wb,save_date,debug):
    if debug is False:
        path = 'C:\Users\kelly.richardson\OneDrive - Imagine Learning Inc\Reports\Efficiency Reports'
        file_name = lead_name+"_"+save_date+"-LEADBOOK.xlsx"
        save_location = os.path.join(path,lead_name,'TEAM E-REPORTS')
        if not os.path.isdir(save_location):
            os.makedirs (save_location)
        save_name = save_location+"/"+file_name
        wb.save(save_name)
    else:
        wb.save('Output/Teacher Books/'+lead_name+"_"+save_date+'.xlsx')



##Do df[[teachername1,teachername2,teachername3]]
##Then set teachername1 as leadname
team_org = {'Jeremy Shock':['Jeremy Shock', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Crystal Boris', 'Melissa Mitchell', 'Cassie Ulisse', 'Laura Gardiner', 'Michelle Amigh', 'Kimberly Stanek'],
    'Rachel Adams':['Rachel Adams', 'Cristen Phillipsen', 'Heather Chilleo', 'Hester Southerland', 'Jamie Weston', 'James Hare', 'Michele  Irwin', 'Juventino Mireles'],
    'Melissa Cox':['Melissa Cox', 'Clifton Dukes', 'Kelly Richardson', 'Veronica Alvarez', 'Nancy Polhemus', 'Kimberly Abrams', 'Stacy Good'],
    'Jill Szafranski':['Salome Saenz', 'Alisa Lynch', 'Gabriela Torres', 'Wendy Bowser', 'Nicole Marsula', 'Donita Farmer', 'Andrea Burkholder', 'Laura Craig', 'Bill Hubert', 'Erin Hrncir'],
    'Kristin Donnelly':['Kristin Donnelly', 'Angel Miller', 'Marcella Parks', 'Sara  Watkins', 'Shannon Stout', 'Lisa Duran', 'Erica Basilone', 'Carol Kish', 'Jennifer Talaski', 'Nicole Knisely'],
    'Caren Glowa':['Caren Glowa', 'Johana Miller', 'Audrey Rogers', 'Cheri Shively', 'Amy Stayduhar', 'Dominique Huffman', 'Meaghan Wright', 'Kathryn Montano', 'Lynae Shepp', 'Anna Bell', 'Jessica Connole']}


"""INPUTS HERE"""
"""Jeremy Shock, Rachel Adams,Melissa Cox, Jill Szafranski,Kristin Donnelly,Caren Glowa, All"""
#Uses Periscope Source and Tabby source to format and make the raw changes sheet in lead book.
lead_name = "All"
periscope = 'e-data_source/e-data_Fall/123_all.csv'
tabby = "e-data_source/e-data_Fall/123_tabby.csv"
create_input(periscope,tabby,lead_name)
print ("Periscope Data Sorted")

#Skip days are used to skip days with bad data, or to only return certain days from a dataset.
skip_days = []
#Used to Conditionally Format the Daily Summary tables
scores = {"Good Day Score": float(.90), "Upper Bound": float(1.25),
'Good Night Score':float(.70)}
#Output Filename that saves file locally. Usually used when testing.
save_date = "12-03-18"
#Used to indicate a end of day for split day function.
end_day_indicator = '12:54 AM'
debug = True



"""Calling Functions"""
wb = load_workbook(filename='Input_EReport.xlsx')
wb_sheet = wb['Sheet1']
wb_sheet.title = 'Raw Changes'
split_sheet_by_days(wb, skip_days, end_day_indicator)
print ("Sheets split by days completed")
call_create_tables(wb)
checks = {'Night Check': False, 'Day Check': False}
checks = define_blocks(wb, checks, scores)
print ("Blocks Found and Daily Tables Created")
df = checks[-1]
checks = checks[0]
data_library = find_non_empty_tables(wb, df)
create_summary_page(wb, data_library, checks)

#This saves leadbook locally in project folder and is used for testing.
wb.save('Output/Fall/'+save_date+'.xlsx')
create_dataframe(data_library,wb, save_date)
#create_books(wb,lead_name, save_date,debug)
#save_leadbook(wb, save_date,debug)




