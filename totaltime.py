import datetime
from datetime import timedelta

jamie = """09:18 AM-11:18 AM
11:48 AM-01:42 PM
02:24 PM-03:48 PM
04:18 PM-05:24 PM
09:12 AM-11:18 AM
11:54 AM-01:48 PM
02:24 PM-03:42 PM
04:18 PM-05:30 PM
09:06 AM-11:12 AM
12:36 PM-01:36 PM
02:18 PM-03:42 PM
04:18 PM-05:36 PM
10:12 AM-11:06 AM
11:42 AM-01:42 PM
02:36 PM-03:42 PM
04:18 PM-05:36 PM
09:30 AM-11:00 AM
11:42 AM-01:42 PM
02:18 PM-03:48 PM
04:24 PM-05:24 PM"""

michele = """08:36 AM-10:36 AM
11:12 AM-11:42 AM
01:42 PM-02:42 PM
04:06 PM-04:54 PM
08:36 AM-10:30 AM
11:06 AM-12:06 PM
01:42 PM-02:42 PM
04:18 PM-05:00 PM
08:36 AM-10:30 AM
11:06 AM-11:18 AM
01:42 PM-02:36 PM
04:06 PM-05:00 PM
09:06 AM-10:42 AM
11:18 AM-12:06 PM
01:42 PM-02:36 PM
04:12 PM-05:00 PM
08:42 AM-10:36 AM
11:12 AM-12:06 PM
01:42 PM-02:42 PM
04:06 PM-04:54 PM"""


hester="""08:30 AM-10:36 AM
11:00 AM-01:06 PM
01:36 PM-03:12 PM
03:36 PM-04:54 PM
09:18 AM-10:42 AM
11:12 AM-01:06 PM
01:36 PM-03:06 PM
03:42 PM-05:00 PM
07:42 PM-08:48 PM
08:36 AM-10:42 AM
11:12 AM-11:36 AM
12:18 PM-01:00 PM
01:30 PM-03:00 PM
03:36 PM-05:06 PM
08:30 AM-10:36 AM
11:06 AM-01:12 PM
01:36 PM-03:06 PM
03:36 PM-05:00 PM
08:30 AM-10:36 AM
11:06 AM-01:12 PM
01:42 PM-03:00 PM
03:30 PM-04:54 PM"""



times ="""08:12 AM-10:06 AM
11:00 AM-12:06 PM
12:42 PM-02:00 PM
08:06 AM-09:30 AM
11:24 AM-12:06 PM
12:42 PM-02:12 PM
08:06 AM-09:00 AM
09:36 AM-10:12 AM
11:00 AM-11:42 AM
09:12 AM-10:00 AM
10:54 AM-12:06 PM
12:42 PM-02:12 PM"""


list_of_times = times.split('\n')
duration = timedelta
for item in list_of_times:
    start = item[:item.index('-')]
    end = item[item.index('-')+1:]
    start = datetime.datetime.strptime(start, '%I:%M %p')
    end = datetime.datetime.strptime(end, '%I:%M %p')
    try:
       sum_diff = sum_diff + end-start
    except:
        sum_diff = end - start
    print sum_diff
secs = sum_diff.total_seconds()
print sum_diff
print secs/60
print sum_diff
