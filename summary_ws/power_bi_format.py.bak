import pandas as pd 
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
    
def create_dataframe(data_dict, wb,save_date,debug):
    #Data_dict is a dictionary of dicrionarys. It's {date:{teacher{night:escore, day:escore}}}
    df = pd.DataFrame({'Teacher':[], 'Team':[], 'Date':[], 'E-Score':[], 'Month_Num':[], 'Weekday':[], 'Weekday_Num':[]})
    for date, value in data_dict.iteritems():
        date_dict = data_dict[date]
        ws = wb[date]
        #Gets date from Worksheet so it includes the year versus just the ws title which does not.
        full_date = ws.cell(row=2, column=6).value
        #Creates a datetime object and then paress it a bunch of different ways to get values
        dt_obj = datetime.datetime.strptime(full_date, '%m/%d/%y %a %I:%M %p')
        dater = dt_obj.strftime('%m/%d/%y')
        weekday = dt_obj.strftime('%m/%d/%y %A')
        month_num = dt_obj.strftime('%m')
        weekday_num = dt_obj.strftime('%w')
        for teacher, data in date_dict.iteritems():
            team = find_team(teacher)

            
            teacher_dict = date_dict[teacher]
            if "  " in teacher:
                teacher = teacher.replace("  "," ")
            df = df.append(({'Teacher':teacher, 'Team':team, 'Date':dater, 
                'E-Score':(teacher_dict['Night Average']),'Day/Night':'Night', 'Month_Num':month_num, 
                'Weekday':weekday, 'Weekday_Num':weekday_num}), ignore_index=True)
            df = df.append(({'Teacher':teacher, 'Team':team, 'Date':dater, 
                'E-Score':(teacher_dict['Day Average']),'Day/Night':'Day', 'Month_Num':month_num, 
                'Weekday':weekday, 'Weekday_Num':weekday_num}), ignore_index=True)

    df = df[['Teacher', 'Team', 'Date', 'E-Score', 'Day/Night', 'Month_Num', 'Weekday', 'Weekday_Num']]
    create_workbook(df,save_date,debug)

def create_workbook(df,save_date,debug):
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    for cell in ws['A'] + ws[1]:
        cell.style = 'Pandas'
    if debug is True:
        wb.save('Output/Fall/'+save_date+'_Management.xlsx')
    else:
        path = "C:\Users\kelly.richardson\OneDrive - Imagine Learning Inc\Reports\Efficiency Reports\Teaching Department"
        file_name = save_date+'_Management.xlsx'
        save_name = os.path.join(path,file_name)
        wb.save(save_name)


def find_team(teacher):
    team_org ={'Jeremy Shock':['Jeremy Shock','Crystal Boris', 'Jamie Weston', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Laura Gardiner', 'Melissa Mitchell', 'Stacy Good', 'Veronica Alvarez'],
    'Rachel Adams':['Rachel Adams', 'Clifton Dukes', 'Heather Chilleo', 'Hester Southerland', 'Juventino Mireles', 'Kelly Richardson', 'Kimberly Stanek', 'Michele  Irwin', 'Michelle Amigh', 'Nancy Polhemus'],
    'Melissa Cox':['Melissa Cox','Emily McKibben', 'Erica De Coste', 'Erin Hrncir', 'Jennifer Talaski', 'Lisa Duran', 'Marcella Parks','Preston Tirey','Erin Spilker'],
    'Sara  Watkins':[ 'Sara  Watkins','Alisa Lynch', 'Andrea Burkholder', 'Bill Hubert', 'Donita Farmer', 'Laura Craig', 'Nicole Marsula', 'Salome Saenz', 'Wendy Bowser'],
    'Kristin Donnelly':['Kristin Donnelly', 'Angel Miller', 'Carol Kish', 'Erica Basilone', 'Euna Pineda', 'Gabriela Torres', 'Jenni Alexander', 'Nicole Knisely', 'Shannon Stout'],
    'Caren Glowa':['Caren Glowa','Amy Stayduhar', 'Audrey Rogers', 'Cheri Shively', 'Jessica Connole', 'Johana Miller', 'Kathryn Montano', 'Lynae Shepp', 'Meaghan Wright','Veraunica Wyatt']}
    for team_lead, teams in team_org.iteritems():
        if teacher in teams:
            names = team_lead.split(" ")
            return names[0]
