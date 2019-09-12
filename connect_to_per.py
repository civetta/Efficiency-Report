import urllib2
import pandas as pd
import datetime
from datetime import timedelta
from connect_to_warehouse import make_connection
import pyodbc


#TEST AFTER PRODUCTS NOT THE BEFORE
#SSMAX IS CALCULATED SLIGHTLY DIFFERENT FIGURE OUT WHY
#NUMBERS ARE SLIGHTLY DIFFERENT, BUT THEY FOLLOW JAIROS CHART TO THE T SO THEY ARE MORE ACCURATE? 
def open_session_closed_data(start_date,end_date):
    #Connects to Periscope to retreieve table
<<<<<<< Updated upstream
    url = "https://app.periscopedata.com/api/think-through-learning/chart/csv/88a7a9a3-169d-833b-ce6c-05de1102841e/487533"
    response = urllib2.urlopen(url)
    df = pd.read_csv(response)
=======
    df = make_connection(start_date,end_date)
>>>>>>> Stashed changes
    #Cleaning up Data from Periscope
    df = df.applymap(str)
    df = df.apply(lambda x: x.str.strip())
    df['date'] = df.stamp + " " + df.end_time
    df =  df[['name','date','completed_sessions']]

    return df

def pivot_df(df):
    #Pivoting and filling na
    df = df.pivot(index='date', columns='name', values = 'completed_sessions')
    df = df.fillna(0)
    df = df.reset_index(drop=False)
    #IDEA: Find first day, then just replace datetime with 7:00AM and fill in columns with 0. Then asfeq will deal with rest.
    #Notes: date = datetime.strptime('26 Sep 2012', '%d %b %Y').replace(hour=7)
    df['date'] = df.date.apply(lambda x: datetime.datetime.strptime(x, '%m-%d-%Y %I:%M %p'))
    print (df['date'])
    df['date'] = df.date.apply(lambda x: x - timedelta(hours=5))
    df = df.set_index('date', drop=True)
    df = df.sort_index()
    cut_date =  df.first_valid_index().replace(hour=7, minute=0)
    mask = (df.index > cut_date)
    df = df.loc[mask]
    first_date = df.first_valid_index()
    df = df.asfreq(freq='420S', fill_value=0)
    df = df.apply(pd.to_numeric, errors='ignore')
    return df

def find_6min_intervals(df, ssmax):
    #Finding average SSMAX for every 6 minute time period, using the date columns as ranges.

    list_of_times =  df.index.values
    ssmax_cols = pd.DataFrame({'date': [list_of_times[0]],'*SSMax': [0]})
    timestamp_loc = 0
    while timestamp_loc < len(list_of_times)-1:
        mask = (ssmax['Per_minute'] > list_of_times[timestamp_loc]) & (ssmax['Per_minute'] <= list_of_times[timestamp_loc+1])
        df3 = ssmax.loc[mask]
        tab = calculate_mean_ssmax(df3, list_of_times[timestamp_loc]) 
        ssmax_cols = ssmax_cols.append({'date' : list_of_times[timestamp_loc+1] , '*SSMax' : tab},ignore_index=True)
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



def get_ssmax(start_date, end_date):
    conn = pyodbc.connect(driver='{SQL Server}', server='VOO1DB2.ILPROD.LOCAL', database='ResearchMarketing', trusted_connection='Yes')
    sql="""with complete as (SELECT
    dateadd(minute,datediff(minute,0,MEASUREMENT_DATE),0) AS Per_Minute,
    CONVERT(date, MEASUREMENT_DATE) AS "DATE_MEASURED",
    DATENAME(MONTH, DATEADD(MONTH,DATEPART(MM,MEASUREMENT_DATE),0)-1) AS "MONTH",
    DATENAME(weekday,MEASUREMENT_DATE)AS DAY_OF_THE_WEEK,
    DATEPART(HH, MEASUREMENT_DATE) AS Hour_of_Day,
    AVG(TEACHERS_ONLINE) as Active_Teachers_Avg,
    AVG(STUDENT_CHATS) as Students_in_Chat_Avg,
    AVG(STUDENTS_QUEUED) as Students_in_Queue_Avg,
    CAST(ROUND(AVG(AVG_SESSIONS_PER_TEACHER),2) AS float(1)) as SS_Avg,
    CAST(ROUND(AVG(MAX_SESSIONS_PER_TEACHER),2) AS float(1)) as SS_Max_Avg
    FROM LIVE_TEACHING_ARCHIVE
    WHERE
    DATEPART(dw,MEASUREMENT_DATE) BETWEEN 2 AND 6
    AND DATEPART(dw,UPLOAD_DATE) BETWEEN 3 AND 7
    AND DATEPART(HH,MEASUREMENT_DATE) BETWEEN 7 AND 24
    GROUP BY
    DATEPART(HH,MEASUREMENT_DATE),
    CONVERT(date, MEASUREMENT_DATE),
    DATENAME(weekday,MEASUREMENT_DATE),
    DATEADD(MONTH,DATEPART(MM,MEASUREMENT_DATE),0)-1,
    dateadd(minute,datediff(minute,0,MEASUREMENT_DATE),0)
    )

    SELECT Per_minute, Active_Teachers_Avg
    , CASE WHEN SS_Max_Avg < 5 then SS_Max_Avg else 5 end as SS_Max_5
    , SS_Avg
    FROM complete
    WHERE Per_Minute between '"""+start_date+"' and '"+end_date+"""' order by Per_Minute desc;"""
    df = pd.read_sql_query(sql, conn)
    df = df[['SS_Max_5','Per_minute']]
    return df


def get_inputs(start_date, end_date):
    ssmax = get_ssmax(start_date, end_date)
    df = open_session_closed_data(start_date, end_date)
    df = pivot_df(df)
    ssmax_cols = find_6min_intervals(df, ssmax)
    result = pd.concat([df, ssmax_cols], axis=1, sort=False)
    result.index = result.index.map(lambda x: x.strftime('%m/%d/%y %a %I:%M %p'))
     all_columns = ['*SSMax', 'Laura Gardiner',  'Caren Glowa', 'Crystal Boris', 'Jamie Weston', 'Kay Plinta-Howard', 'Marcella Parks', 'Melissa Mitchell', 'Michelle Amigh', 'Stacy Good',  
'Rachel Adams', 'Clifton Dukes', 'Heather Chilleo', 'Hester Southerland', 'Kelly-Anne Heyden', 'Kimberly Stanek', 'Michele Irwin', 'Nancy Polhemus', 'Juventino Mireles',  
'Melissa Cox', 'Andrew Lowe', 'Emily McKibben', 'Erica DeCosta', 'Erin Hrncir', 'Erin Spiker', 'Jennifer Talaski', 'Julie Horne', 'Lisa Duran', 'Preston Tirey',   
'Sara Watkins', 'Alisa Lynch', 'Andrea Burkholder', 'Angela Miller', 'Bill Hubert', 'Donita Spencer', 'Jessica Connole', 'Laura Craig', 'Nicole Marsula', 'Rachel Romana', 'Veronica Alvarez', 'Wendy Bowser', 
'Kristin Donnelly', 'Carol Kish', 'Erica Basilone', 'Euna Pin', 'Hannah Beus', 'Jenni Alexander', 'Jessica Throolin', 'Natasha Andorful', 'Nicole Knisely', 'Shannon Stout', 
'Gabriela Torres', 'Amy Stayduhar', 'Audrey Rogers', 'Cheri Shively', 'Kathryn Montano', 'Karen Henderson', 'Lynae Shepp', 'Johana Miller', 'Meaghan Wright', 'Veronica Wyatt']
    result = result.reindex(columns=all_columns, fill_value=0)
    return result
