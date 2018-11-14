import pandas as pd 
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
    
def create_dataframe(data_dict, wb):
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
            df = df.append(({'Teacher':teacher, 'Team':team, 'Date':dater, 
                'E-Score':(teacher_dict['Night Average']),'Day/Night':'Night', 'Month_Num':month_num, 
                'Weekday':weekday, 'Weekday_Num':weekday_num}), ignore_index=True)
            df = df.append(({'Teacher':teacher, 'Team':team, 'Date':dater, 
                'E-Score':(teacher_dict['Day Average']),'Day/Night':'Day', 'Month_Num':month_num, 
                'Weekday':weekday, 'Weekday_Num':weekday_num}), ignore_index=True)

    df = df[['Teacher', 'Team', 'Date', 'E-Score', 'Day/Night', 'Month_Num', 'Weekday', 'Weekday_Num']]
    create_workbook(df)

def create_workbook(df):
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    for cell in ws['A'] + ws[1]:
        cell.style = 'Pandas'
    wb.save('Output/Fall/11-5_Management.xlsx')


def find_team(teacher):
    team_org = {'Jeremy Shock':['Jeremy Shock', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Crystal Boris', 'Melissa Mitchell', 'Cassie Ulisse', 'Laura Gardiner', 'Michelle Amigh', 'Kimberly Stanek'],
    'Rachel Adams':['Rachel Adams', 'Cristen Phillipsen', 'Heather Chilleo', 'Hester Southerland', 'Jamie Weston', 'James Hare', 'Michele  Irwin', 'Juventino Mireles'],
    'Melissa Cox':['Melissa Cox', 'Clifton Dukes', 'Kelly Richardson', 'Veronica Alvarez', 'Nancy Polhemus', 'Kimberly Abrams', 'Stacy Good'],
    'Jill Szafranski':['Salome Saenz', 'Alisa Lynch', 'Gabriela Torres', 'Wendy Bowser', 'Nicole Marsula', 'Donita Farmer', 'Andrea Burkholder', 'Laura Craig', 'Bill Hubert', 'Erin Hrncir'],
    'Kristin Donnelly':['Kristin Donnelly', 'Angel Miller', 'Marcella Parks', 'Sara  Watkins', 'Shannon Stout', 'Lisa Duran', 'Erica Basilone', 'Carol Kish', 'Jennifer Talaski', 'Nicole Knisely'],
    'Caren Glowa':['Caren Glowa', 'Johana Miller', 'Audrey Rogers', 'Cheri Shively', 'Amy Stayduhar', 'Dominique Huffman', 'Meaghan Wright', 'Kathryn Montano', 'Lynae Shepp', 'Anna Bell', 'Jessica Connole']}
    for team_lead, teams in team_org.iteritems():
        if teacher in teams:
            names = team_lead.split(" ")
            return names[0]
