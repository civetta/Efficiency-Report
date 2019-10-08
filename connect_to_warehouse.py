import psycopg2
import pandas as pd
import datetime


def make_connection(start_date,end_date):
    conn = psycopg2.connect("dbname='warehouse' user='kellyrichardson'  password='8b3c9XFGLj3FiSnQvzfJx' host='im-warehouse-prod.cfozmy0xza77.us-west-2.rds.amazonaws.com'")
    sql = """With Sub_Table as ( Select Name,
      year_number,
      month_number,
      month_name,
      stamp,
      Day_of_Week,
        to_char(Datetime,'HH:MI AM') ||' - '|| to_char(Datetime,'HH:MI AM')  as time_of_the_day,
        to_char(Datetime+INTERVAL'1           minute','HH:MI AM')  as end_time,
        Count(*) as Completed_Sessions,
        Team
    From
    (SELECT 
    teachers.name,
    dates.year_number,
    dates.month_number,
    dates.month_name,
    dates.date as stamp,      
    dates.day_of_week,
    times.hour,
    TIMESTAMP 'epoch' + 60*floor(date_part('epoch', CAST(dates.date as date) + CAST(times.time as time))/(60))        *INTERVAL '1 second' as Datetime
    --Lauras's Team
    ,CASE WHEN teachers.id IN (
        152964	--Caren Glowa
        ,270102	--Crystal Boris
        ,548352	--Jamie Weston
        ,5955	--Kay Plinta-Howard
        ,725743	--Marcella Parks
        ,5957	--Melissa Mitchell
        ,725733	--Michelle Amigh
        ,592154	--Stacy Good
        ,723678	--Laura Gardiner
    ) THEN 'Laura' 
            
    --Rachel's Team
    WHEN teachers.id IN (
    5962	--Rachel Adams
    ,901687 --Clifton Dukes
    ,5960	--Heather Chilleo
    ,273045	--Hester Southerland
    ,205305	--Kelly-Anne Heyden
    ,723676	--Kimberly Stanek
    ,555127	--Michele Irwin
    ,553281	--Nancy Polhemus
    ,5966	--Juventino Mireles
        ) THEN 'Rachel'


    --Melissa's Team
    WHEN teachers.id IN (
        8444	--Melissa Cox
        ,1027651	--Andrew Lowe
    ,983167 -- Emily McKibben
    ,985473 --Erica DeCosta
    ,594225	--Erin Hrncir
        ,997469 --Erin Spiker
    ,555566	--Jennifer Talaski
    ,1027654	--Julie Horne
    ,559642	--Lisa Duran
    ,993319 --Preston Tirey
    
        ) THEN 'Melissa'

    --Sara's Team
    WHEN teachers.id IN (
    548353	--Sara Watkins
    ,150843	--Alisa Lynch
    ,725737	--Andrea Burkholder
    ,555257	--Angela Miller
    ,274007	--Bill Hubert
    ,188078	--Donita Spencer
    ,896147 --Jessica Connole
    ,587414	--Laura Craig
    ,40394	--Nicole Marsula
    ,1028752 --Rachel Romana
    ,555565	--Veronica Alvarez
    ,5952	--Wendy Bowser
        ) THEN 'Sara'

    --Kristin's Team
    WHEN teachers.id IN (
        5958	--Kristin Donnelly
    
        ,280470	--Carol Kish
    ,555126	--Erica Basilone
    ,984490 --Euna Pin
    ,1029083	--Hannah Beus
    ,982757 --Jenni Alexander
    
    ,1027653	--Jessica Throolin
    ,1027652	--Natasha W/
    ,555128	--Nicole Knisely
        ,6516	--Shannon Stout
        ) THEN 'Kristin'

    --Gabby's Team
    WHEN teachers.id IN (
    164866	--Gabriela Torres
    ,6515	--Amy Stayduhar
    ,5965	--Audrey Rogers
    ,262061	--Cheri Shively
    ,596811	--Kathryn Montano
    ,1028751 --Karen Henderson
    ,725746	--Lynae Shepp
    ,553279	--Johana Miller
    ,278244	--Meaghan Wright
    ,997470 --Veronica Wyatt
        ) THEN 'Gabby'  
            
    ELSE 'n/a'
    END AS Team


    FROM live_help_facts
    left join teachers on live_help_facts.teacher_id = teachers.id
    left join dates on live_help_facts.utc_completed_date_id = dates.id
    left join times on live_help_facts.utc_completed_time_id = times.id
    WHERE email LIKE '%liveteacher%'
    and transcript  is not NULL
    and Cast(dates.date as Date) between '"""+start_date+"""' and '"""+end_date+"""'
    ) as temp_table
                    
    --WHERE [Teacher_Name=Live_Teacher_Names]
    --and [Team=Teacher_Team]
    --and [Datetime=daterange]
    GROUP BY temp_table.datetime, 1,2,3,4,5,6,7,10
    ORDER BY Datetime DESC
    LIMIT 50000)




    select Name,Team, stamp, end_time, Completed_Sessions from Sub_Table"""
    df = pd.read_sql_query(sql,conn)
    
    df['date'] = df['stamp'] + " " + df['end_time']
    
    df['date'] = df.date.apply(lambda x: datetime.datetime.strptime(x, "%m-%d-%Y %I:%M %p"))
    df = daylights(df)
    start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d")
    df = df[(df['date'] > start_date)]
    df['date'] = df.date.apply(lambda x: x.strftime('%m/%d/%y %I:%M %p'))
    df =  df[['name','date','completed_sessions']]
    
    return df


def daylights(df):        
    #df['date'] = df.end_time.apply(lambda x: datetime.datetime.strptime(x, "%I:%M %p"))
    clifton_subset = df.name == 'Clifton Dukes'
    clifton = df[clifton_subset]
    df.to_csv('0.csv')
    df['date'] = df['date'] - pd.Timedelta(hours=4)
    clifton_subset = df.name == 'Clifton Dukes'
    clifton = df[clifton_subset]
    df.to_csv('1.csv')
    #df['date'] = df.end_time.apply(lambda x: x.strftime("%I:%M %p"))
    return df