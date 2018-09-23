import pandas as pd 
import numpy as np
import datetime
import math
from datetime import timedelta
from dateutil import tz

def create_input(periscope,tabby):
    import warnings
    warnings.filterwarnings("ignore")
    pd.set_option('mode.chained_assignment', None)
    #Input Variables
    df = pd.read_csv(periscope)
    Tabby = pd.read_csv(tabby)
    Tabby = Tabby.dropna()
    Tabby = organize_Tabby(Tabby)
    #Creates an "end_timestamp" column, with correct format
    df['end_timestamp'] = df['session_ended'].map(create_timestamp)
    df['end_date'] = df['end_timestamp'].map(lambda x: x.date())
    #Sets end_timestamps as index and sorts them
    df = df.set_index('end_timestamp')
    df = df.sort_index()
    #Creates an array of dfs, where each df is 1 teacher.
    seperate_days(df,Tabby)

def organize_Tabby(Tabby):
    new_Tabby = Tabby[['Per_minute','SS_Max_Avg']]
    new_Tabby.columns = ['Stamp','Tabby']

    
    new_Tabby['timestamp']=new_Tabby['Stamp'].map(lambda x: datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S.%f"))
    new_Tabby['date'] = new_Tabby['timestamp'].map(lambda x: x.strftime('%Y-%m-%d'))
    new_Tabby['DateStamp'] = new_Tabby['date'].map(lambda x: datetime.datetime.strptime(x, "%Y-%m-%d"))
    new_Tabby = new_Tabby[['timestamp','DateStamp','Tabby']]
    new_Tabby = new_Tabby.set_index('timestamp')
    new_Tabby = new_Tabby.sort_index()
    return new_Tabby


def create_timestamp(x):
    #Removes milisecond from timestamp and formats in datetime.
    try:
        x = x[:x.index('.')]
    except ValueError:
        pass
    x = datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S")
    return x

def seperate_days(df,Tabby):
    #Finds First and Laste day in df, and then iterates through them.
    first_day = df.end_date.values[0]
    last_day = df.end_date.values[-1]
    delta = datetime.timedelta(days=1)
    sessions_ended = pd.DataFrame()
    unique_name = df.teacher_name.unique()
    unique_name.sort()
    week_df = pd.DataFrame()

    while first_day <= last_day:
        current_day_df = df[(df['end_date']==first_day)]
        start = current_day_df.index.values[0]
        start = pd.Timestamp(start)	        
        start = start.replace(hour=0, minute=0,second=0)	  
        end = start.replace(hour=23, minute=54, second=0)
        day_Tabby = Tabby[(Tabby['DateStamp']==first_day)]
        Tabby_col = organize(day_Tabby,start,end,'Tabby')
        
        day_df=pd.DataFrame()
        day_df['Tabby'] = Tabby_col['Tabby']
        for teacher_name in unique_name:
            #Creates a df for each teacher on each day
            teacher_per_day = current_day_df.loc[current_day_df['teacher_name'] == teacher_name]
            #Finds session closed every 6 minutes
            new_column = organize(teacher_per_day,start,end,teacher_name)
            day_df[teacher_name] = new_column[teacher_name]
            
        week_df = week_df.append(day_df)
        first_day += delta
    week_df.set_index(['Tabby'])
    week_df = week_df.fillna(0)
    week_df.index = week_df.index.map(fix_timestamp)
    print week_df.index.values[0]
    print week_df.columns
    
    week_df.to_csv('Week.csv')
    writer = pd.ExcelWriter('Input_EReport.xlsx')
    week_df.to_excel(writer, index = True)
    writer.save()

def fix_timestamp(x):
    dt_obj = datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S")
    dt_obj = dt_obj.strftime('%m/%d/%y %a %I:%M %p')
    return dt_obj


#IS TABBY SHIFTED DOWN 1 CELL!!!???????


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
        
        if column_name == 'Tabby':
            tab = df2['Tabby'].mean()
            tab = math.ceil(tab)
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

