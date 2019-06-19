import datetime
from datetime import timedelta

times="""09:48 AM-11:00 AM
12:12 PM-01:36 PM
02:48 PM-04:30 PM
06:12 PM-06:36 PM
09:48 AM-11:36 AM
12:18 PM-01:36 PM
02:54 PM-04:30 PM
06:06 PM-06:30 PM
09:42 AM-11:30 AM
12:12 PM-01:30 PM
02:48 PM-04:30 PM
06:12 PM-06:30 PM
10:06 AM-11:36 AM
12:18 PM-01:36 PM
02:48 PM-04:12 PM
06:12 PM-06:30 PM
09:42 AM-11:36 AM
12:30 PM-01:36 PM
03:06 PM-04:30 PM
06:06 PM-06:18 PM"""


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
