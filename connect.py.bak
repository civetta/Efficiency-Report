import pyodbc
import pandas as pd

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
WHERE Per_Minute between """+start_date+ " and "+end_date+"""
 order by Per_Minute desc;"""
df = pd.read_sql_query(sql, conn)
print df