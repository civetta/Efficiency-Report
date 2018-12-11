import pandas as pd
from datetime import datetime

def alter_df(gspread):
    df = gspread[['date','time','ssmax']]
    df = df.dropna()
    df['Per_minute'] = df['date'] + " "+df['time']
    df['Per_minute'] = df.Per_minute.apply(lambda x: str(x))
    df['Per_minute'] = df.Per_minute.apply(lambda x: datetime.strptime(x, '%m/%d/%Y %I:%M:%S %p'))
    df['Per_minute'] = df.Per_minute.apply(lambda x: datetime.strftime(x, '%Y-%m-%d %H:%M:%S.%f'))
    print df[df['ssmax']>5.0]
    df['SS_Max_5'] = df.ssmax.apply(lambda x: float(5.0) if float(x) > float(5.0) else float(x))
    print df[df['SS_Max_5']>5.0]
    #df2 = pd.read_csv('e-data_source/e-data_Fall/1126_tabby2.csv')
    df.to_csv('e-data_source/e-data_Fall/1126_tabby52.csv')


gspread = pd.read_csv('e-data_source/e-data_Fall/1126_tabby5.csv')
alter_df(gspread)
