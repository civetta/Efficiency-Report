from openpyxl import load_workbook
from periscope_source import create_input
from split_days import split_sheet_by_days
from daily_ws.create_tables import call_create_tables
from daily_ws.mark_blocks import define_blocks
from summary_ws.calculate_daily_escore import find_non_empty_tables
from summary_ws.efficiency_score_summary import create_summary_page
from teacherbooks.create_teacher_books import create_books
from summary_ws.power_bi_format import create_dataframe
from connect_to_per import get_inputs
from datetime import datetime
import os
import pandas as pd
import warnings
warnings.filterwarnings("ignore")



def save_leadbook(wb,save_date,debug, lead_name):
    if debug is False:
        path = 'C:\Users\kelly.richardson\OneDrive - Imagine Learning Inc\Reports\Efficiency Reports'
        file_name = lead_name+"_"+save_date+"-LEADBOOK.xlsx"
        save_location = os.path.join(path,lead_name,'TEAM E-REPORTS')
        if not os.path.isdir(save_location):
            os.makedirs (save_location)
        save_name = save_location+"/"+file_name
        wb.save(save_name)
    else:
        wb.save('Output/Teacher Books/LEADBOOK_'+lead_name+"_"+save_date+'.xlsx')
        

#Skip days are used to skip days with bad data, or to only return certain days from a dataset.
skip_days = []
#Used to Conditionally Format the Daily Summary tables
scores = {"Good Day Score": float(.90), "Upper Bound": float(1.25),
'Good Night Score':float(.70)}
#Output Filename that saves file locally. Usually used when testing.

save_date = "12-17-18"
#Used to indicate a end of day for split day function.
end_day_indicator = '12:54 AM'
debug = False

##Do df[[teachername1,teachername2,teachername3]]
##Then set teachername1 as leadname
team_org = [['Jeremy Shock','*SSMax','Jeremy Shock', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Crystal Boris', 'Melissa Mitchell', 'Cassie Ulisse', 'Laura Gardiner', 'Michelle Amigh', 'Kimberly Stanek'],
    ['Rachel Adams','*SSMax','Rachel Adams', 'Cristen Phillipsen', 'Heather Chilleo', 'Hester Southerland', 'Jamie Weston', 'Michele  Irwin', 'Juventino Mireles'],
    ['Melissa Cox','*SSMax','Melissa Cox', 'Clifton Dukes', 'Kelly Richardson', 'Veronica Alvarez', 'Nancy Polhemus', 'Kimberly Abrams', 'Stacy Good'],
    ['Jill Szafranski','*SSMax','Salome Saenz', 'Alisa Lynch', 'Gabriela Torres', 'Wendy Bowser', 'Nicole Marsula', 'Donita Farmer', 'Andrea Burkholder', 'Laura Craig', 'Bill Hubert', 'Erin Hrncir'],
    ['Kristin Donnelly','*SSMax','Kristin Donnelly', 'Angel Miller', 'Marcella Parks', 'Sara  Watkins', 'Shannon Stout', 'Lisa Duran', 'Erica Basilone', 'Carol Kish', 'Jennifer Talaski', 'Nicole Knisely'],
    ['Caren Glowa','*SSMax','Caren Glowa', 'Johana Miller', 'Audrey Rogers', 'Cheri Shively', 'Amy Stayduhar', 'Dominique Huffman', 'Meaghan Wright', 'Kathryn Montano', 'Lynae Shepp', 'Anna Bell', 'Jessica Connole'],
    ['All','*SSMax','Jeremy Shock', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Crystal Boris', 'Melissa Mitchell', 'Cassie Ulisse', 'Laura Gardiner', 'Michelle Amigh', 'Kimberly Stanek',
    'Rachel Adams', 'Cristen Phillipsen', 'Heather Chilleo', 'Hester Southerland', 'Jamie Weston', 'Michele  Irwin', 'Juventino Mireles',
    'Melissa Cox', 'Clifton Dukes', 'Kelly Richardson', 'Veronica Alvarez', 'Nancy Polhemus', "Kimberly Abrams", 'Stacy Good',
     'Salome Saenz', 'Alisa Lynch', 'Gabriela Torres', 'Wendy Bowser', 'Nicole Marsula', 'Donita Farmer', 'Andrea Burkholder', 'Laura Craig', 'Bill Hubert', 'Erin Hrncir',
    'Kristin Donnelly', 'Angel Miller', 'Marcella Parks', 'Sara  Watkins', 'Shannon Stout', 'Lisa Duran', 'Erica Basilone', 'Carol Kish', 'Jennifer Talaski', 'Nicole Knisely',
    'Caren Glowa', 'Johana Miller', 'Audrey Rogers', 'Cheri Shively', 'Amy Stayduhar', 'Dominique Huffman', 'Meaghan Wright', 'Kathryn Montano', 'Lynae Shepp', 'Anna Bell', 'Jessica Connole']]

lead_name = "All"
start_date ='2018-12-03'
end_date = '2018-12-08'
week_df = get_inputs(start_date, end_date)
week_df = week_df.sort_index(axis=1)
writer = pd.ExcelWriter(start_date+'_input.xlsx')
week_df.to_csv('new_input_123.csv')
week_df.to_excel(writer, index = True)
writer.save()

for team in team_org:
    lead_name = team[0]
    print lead_name
    
    #Get team_df subset, clean it up, and save it as an excel file.
    team_sliced = team[1:]
    team_df = week_df[team_sliced]
    team_df = team_df.sort_index(axis=1)
    team_df.rename(columns={'*SSMax':'SSMax'}, inplace=True)
    writer = pd.ExcelWriter(lead_name+'_input.xlsx')
    team_df.to_excel(writer, index = True)
    #writer.save()

    #Organize Excel File
    wb = load_workbook(filename=lead_name+'_input.xlsx')
    wb_sheet = wb['Sheet1']
    wb_sheet.title = 'Raw Changes'


    split_sheet_by_days(wb, skip_days, end_day_indicator)
    call_create_tables(wb)
    checks = {'Night Check': False, 'Day Check': False}
    checks = define_blocks(wb, checks, scores)
    """df = checks[-1]
    checks = checks[0]
    data_library = find_non_empty_tables(wb, df)
    create_summary_page(wb, data_library, checks)
    wb.save(lead_name+'_leadbook.xlsx')
    if lead_name =='All':
        create_dataframe(data_library,wb, save_date)
        save_leadbook(wb, save_date,debug, lead_name)
    else:
        create_books(wb,lead_name, save_date,debug)
        save_leadbook(wb, save_date,debug, lead_name)"""








