import pandas as pd
from datetime import datetime

def alter_df(gspread):
    df = gspread[['date','time','ssmax']]
    df = df.dropna()
    df['Per_minute'] = df['date'] + " "+df['time']
    print (df)
    df['Per_minute'] = df.Per_minute.apply(lambda x: str(x))
    df['Per_minute'] = df.Per_minute.apply(lambda x: datetime.strptime(x, '%m/%d/%Y %I:%M:%S %p'))
    df['Per_minute'] = df.Per_minute.apply(lambda x: datetime.strftime(x, '%Y-%m-%d %H:%M:%S.%f'))
    df['Per_minute'] = df.Per_minute.apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S.%f'))
    #df2 = pd.read_csv('e-data_source/e-data_Fall/1126_tabby2.csv')
    df.to_csv('fallsource.csv')


gspread = pd.read_csv('falltabby.csv')
alter_df(gspread)
