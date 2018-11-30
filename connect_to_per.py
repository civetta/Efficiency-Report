import urllib2
import pandas as pd
from datetime import datetime

url = "https://app.periscopedata.com/api/think-through-learning/chart/csv/88a7a9a3-169d-833b-ce6c-05de1102841e/487533"
response = urllib2.urlopen(url)

df = pd.read_csv(response)
df = df.applymap(str)
df = df.apply(lambda x: x.str.strip())
df['date'] = df.full_day + " " + df.end_time

df['date'] = df.date.apply(lambda x: datetime.strptime(x, "%m/%d/%y %A %H:%M %p" ))
print df
"""datetime.strptime
df =  df[['team','teacher_name','time_of_the_day','completed_sessions','full_date']]
print df
df.rename(columns={'time_of_the_day':'Date'}, inplace=True)
#Seperate By Team Here If Nessecary
unique_dates = df.Date.unique()
names = df.teacher_name.unique()
new_df = pd.DataFrame(columns=names)
new_df['Date'] = unique_dates
new_df = new_df.set_index('Date')
teacher_df = pd.DataFrame()

for date in unique_dates:
    sub_df = df[df['Date']==date]
    
    sub_df = sub_df[['teacher_name','completed_sessions']].T
    sub_df.columns = sub_df.iloc[0]
    sub_df = sub_df[1:] 
    sub_df ['Date'] = date
    #sub_df = sub_df.reset_index()
    teacher_df.append(sub_df)"""
    