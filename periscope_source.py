import pandas as pd 
import numpy as np
import datetime
import math
from datetime import timedelta
from dateutil import tz

def create_input(periscope,SSMax,lead_name):
    import warnings
    warnings.filterwarnings("ignore")
    pd.set_option('mode.chained_assignment', None)
    #Input Variables
    df = pd.read_csv(periscope)
    df = df.dropna()
    SSMax = pd.read_csv(SSMax)
    SSMax = SSMax.dropna()
    SSMax = organize_SSMax(SSMax)
    #Creates an "end_timestamp" column, with correct format
    df['end_timestamp'] = df['session_ended'].map(create_timestamp)
    df['end_date'] = df['end_timestamp'].map(lambda x: x.date())
    #Sets end_timestamps as index and sorts them
    df = df.set_index('end_timestamp')
    df = df.sort_index()
    #Creates an array of dfs, where each df is 1 teacher.
    seperate_days(df,SSMax,lead_name)


def organize_SSMax(SSMax):
    new_SSMax = SSMax[['Per_minute','SS_Max_5']]
    new_SSMax.columns = ['Stamp','SSMax']

    
    new_SSMax['timestamp']=new_SSMax['Stamp'].map(lambda x: datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S.%f"))
    new_SSMax['date'] = new_SSMax['timestamp'].map(lambda x: x.strftime('%Y-%m-%d'))
    new_SSMax['DateStamp'] = new_SSMax['date'].map(lambda x: datetime.datetime.strptime(x, "%Y-%m-%d"))
    new_SSMax = new_SSMax[['timestamp','DateStamp','SSMax']]
    new_SSMax = new_SSMax.set_index('timestamp')
    new_SSMax = new_SSMax.sort_index()
    return new_SSMax


def create_timestamp(x):
    #Removes milisecond from timestamp and formats in datetime.
    try:
        x = x[:x.index('.')]
    except ValueError:
        pass
    x = datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S")
    return x




















def seperate_days(df,SSMax,lead_name):
    #Finds First and Laste day in df, and then iterates through them.
    first_day = df.end_date.values[0]
    last_day = df.end_date.values[-1]
    delta = datetime.timedelta(days=1)
    sessions_ended = pd.DataFrame()
    unique_name = df.name.unique()
    unique_name.sort()
    week_df = pd.DataFrame()

    while first_day <= last_day:
        current_day_df = df[(df['end_date']==first_day)]
        start = current_day_df.index.values[0]
        start = pd.Timestamp(start)	        
        start = start.replace(hour=7, minute=30,second=0)
        end = start + timedelta(days=1)	  
        end = end.replace(hour=1, minute=00, second=0)
        day_SSMax = SSMax[(SSMax['DateStamp']==first_day)]
        SSMax_col = organize(day_SSMax,start,end,'SSMax')
        
        day_df=pd.DataFrame()
        day_df['SSMax'] = SSMax_col['SSMax']
        current_day_df.to_csv('testing_Between_time.csv')
        for teacher_name in unique_name:
            #Creates a df for each teacher on each day
            teacher_per_day = current_day_df.loc[current_day_df['name'] == teacher_name]
            #Finds session closed every 6 minutes
            new_column = organize(teacher_per_day,start,end,teacher_name)
            day_df[teacher_name] = new_column[teacher_name]
            
        week_df = week_df.append(day_df)
        first_day += delta
    
    
    week_df = fill_in_missing_teachers(week_df,lead_name)
    week_df = week_df.fillna(0)
    week_df.index = week_df.index.map(fix_timestamp)
    
    week_df.to_csv('Week.csv')
    writer = pd.ExcelWriter('Input_EReport.xlsx')
    week_df.to_excel(writer, index = True)
    writer.save()

def fix_timestamp(x):
    dt_obj = datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S")
    dt_obj = dt_obj.strftime('%m/%d/%y %a %I:%M %p')
    return dt_obj

def fill_in_missing_teachers(week_df,lead_name):
    team_org = {'Jeremy Shock':['Jeremy Shock', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Crystal Boris', 'Melissa Mitchell', 'Cassie Ulisse', 'Laura Gardiner', 'Michelle Amigh', 'Kimberly Stanek'],
    'Rachel Adams':['Rachel Adams', 'Cristen Phillipsen', 'Heather Chilleo', 'Hester Southerland', 'Jamie Weston', 'James Hare', 'Michele  Irwin', 'Juventino Mireles'],
    'Melissa Cox':['Melissa Cox', 'Clifton Dukes', 'Kelly Richardson', 'Veronica Alvarez', 'Nancy Polhemus', 'Kimberly Abrams', 'Stacy Good'],
    'Jill Szafranski':['Salome Saenz', 'Alisa Lynch', 'Gabriela Torres', 'Wendy Bowser', 'Nicole Marsula', 'Donita Farmer', 'Andrea Burkholder', 'Laura Craig', 'Bill Hubert', 'Erin Hrncir'],
    'Kristin Donnelly':['Kristin Donnelly', 'Angel Miller', 'Marcella Parks', 'Sara  Watkins', 'Shannon Stout', 'Lisa Duran', 'Erica Basilone', 'Carol Kish', 'Jennifer Talaski', 'Nicole Knisely', 'Desiree Sowards'],
    'Caren Glowa':['Caren Glowa', 'Johana Miller', 'Audrey Rogers', 'Cheri Shively', 'Amy Stayduhar', 'Dominique Huffman', 'Meaghan Wright', 'Kathryn Montano', 'Lynae Shepp', 'Anna Bell', 'Jessica Connole'],
    'All':['Jeremy Shock', 'Jennifer Gilmore', 'Kay Plinta-Howard', 'Crystal Boris', 'Melissa Mitchell', 'Cassie Ulisse', 'Laura Gardiner', 'Michelle Amigh', 'Kimberly Stanek',
    'Rachel Adams', 'Cristen Phillipsen', 'Heather Chilleo', 'Hester Southerland', 'Jamie Weston', 'James Hare', 'Michele  Irwin', 'Juventino Mireles',
    'Melissa Cox', 'Clifton Dukes', 'Kelly Richardson', 'Veronica Alvarez', 'Nancy Polhemus', "Kimberly Abrams", 'Stacy Good',
     'Salome Saenz', 'Alisa Lynch', 'Gabriela Torres', 'Wendy Bowser', 'Nicole Marsula', 'Donita Farmer', 'Andrea Burkholder', 'Laura Craig', 'Bill Hubert', 'Erin Hrncir',
    'Kristin Donnelly', 'Angel Miller', 'Marcella Parks', 'Sara  Watkins', 'Shannon Stout', 'Lisa Duran', 'Erica Basilone', 'Carol Kish', 'Jennifer Talaski', 'Nicole Knisely', 'Desiree Sowards',
    'Caren Glowa', 'Johana Miller', 'Audrey Rogers', 'Cheri Shively', 'Amy Stayduhar', 'Dominique Huffman', 'Meaghan Wright', 'Kathryn Montano', 'Lynae Shepp', 'Anna Bell', 'Jessica Connole']}
    team_df = pd.DataFrame.from_dict(team_org,orient='index')
    team_df = team_df.T
    team_col =  team_df[lead_name].dropna()
    columns= week_df.columns.values
    team = team_col.values
    difference = list(set(team) - set(columns))
    for teacher in difference:
        week_df[teacher] = 0
    return week_df


def organize (df,start,end,column_name):
    sessions_ended = {}
    while start != end:
        range1 = str(start)
        range1 = range1[range1.index(" ")+1:]
        range2 = str(start + datetime.timedelta(minutes=6))
        stamprange2 = range2
        range2 = range2[range2.index(" ")+1:]
        num = []
        df2 = df.between_time(range1, range2,include_start=False, include_end=True)
        
        if column_name == 'SSMax':
            tab = df2['SSMax'].mean()
            tab = round(float(tab),2)
            if tab > 5:
                tab = 5
            num.append(tab)
        else:
            num.append(df2.shape[0])
        sessions_ended.update({stamprange2:num})
        start = start + datetime.timedelta(minutes=6)  
    sessions_ended = pd.DataFrame(sessions_ended)
    #Transposes sessions_ended so timestamps are rows, not columns.
    sessions_ended = sessions_ended.T
    #Sets to columns to alphabetical teacher_name
    #Sets timestamp as index
    sessions_ended.index.names = ['Date']
    sessions_ended.columns=[column_name]
    return sessions_ended
