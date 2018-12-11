import urllib2
import pandas as pd
import datetime
from datetime import timedelta

#NOTE TO SELF FIGURE OUT WHY IN BETWEEN TIME ISN'T WORKING


def open_session_closed_data():
    #Connects to Periscope to retreieve table
    url = "https://app.periscopedata.com/api/think-through-learning/chart/csv/88a7a9a3-169d-833b-ce6c-05de1102841e/487533"
    response = urllib2.urlopen(url)
    df = pd.read_csv(response)

    #Cleaning up Data from Periscope
    df = df.applymap(str)
    df = df.apply(lambda x: x.str.strip())
    df['date'] = df.full_day + " " + df.end_time
    df =  df[['teacher_name','date','completed_sessions']]
    return df

def pivot_df(df):
    #Pivoting and filling na
    df = df.pivot(index='date', columns='teacher_name', values = 'completed_sessions')
    df = df.fillna(0)
    df = df.reset_index(drop=False)
    df['date'] = df.date.apply(lambda x: datetime.datetime.strptime(x, '%m/%d/%y %I:%M %p'))
    df = df.set_index('date', drop=True)
    df = df.asfreq(freq='360S', fill_value=0)
    df.index = df.index.map(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))
    return df

def clean_ssmax_data(ssmax):
    #Organize SSMax Input
    ssmax = ssmax[['SS_Max_5','Per_minute']]
    ssmax['Per_minute'] = ssmax.Per_minute.apply(lambda x: datetime.datetime.strptime(x, "%Y-%m-%d %H:%M:%S.%f"))
    #ssmax['Per_minute'] = ssmax.Per_minute.apply(lambda x: x.strftime("%Y-%m-%d %H:%M:%S"))
    ssmax.set_index('Per_minute', drop=True, inplace=True)
    return ssmax

def find_6min_intervals(df, ssmax):
    #Finding average SSMAX for every 6 minute time period, using the date columns as ranges.
    #df1 = pd.DataFrame([['a', 1], ['b', 2]],
                   #columns=['letter', 'number'])
    
    list_of_times =  df.index.values
    ssmax_times = ssmax.index.values
    print list_of_times[0:5]
    print ssmax_times[0:5]
    ssmax_cols = pd.DataFrame({'date': [list_of_times[0]],'SSMax': [0]})
    timestamp_loc = 0
    while timestamp_loc < len(list_of_times)-1:
        
        range1 = list_of_times[timestamp_loc]
        range2 = list_of_times[timestamp_loc+1]
        df3 = ssmax.between_time(range1, range2,include_start=False, include_end=True)
        tab = calculate_mean_ssmax(df3, range1) 
        ssmax_cols = ssmax_cols.append({'date' : range2 , 'SSMax' : tab},ignore_index=True)
        timestamp_loc = timestamp_loc+ 1
    ssmax_cols = ssmax_cols.set_index('date', drop=True)
    return ssmax_cols

def calculate_mean_ssmax(df3, range1):
    if len(df3.index) != 0:
        if df3.index.values[0] == range1:
            df3 = df3.iloc[1:]
        tab = df3['SS_Max_5'].mean()
        tab = round(float(tab),2)
    else:
        tab = 0
    return tab

#SSmax will be called extrernally in future
ssmax = pd.read_csv('e-data_source/e-data_Fall/123_tabby.csv')

df = open_session_closed_data()
df = pivot_df(df)
ssmax = clean_ssmax_data(ssmax)
ssmax_cols = find_6min_intervals(df, ssmax)
#print ssmax_cols.iloc[300:400]
result = pd.concat([df, ssmax_cols], axis=1, sort=False)
#print result.iloc[300:400]
