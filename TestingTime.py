"""
Input: Two string of time. Example: 03/21/18 Thu 8:27 AM and 03/21/18 Thu 9:48 AM.
Output: A string with just the hours (example: 8:27 AM - 9:48 AM").
Times is marked with an asteriks if one of the times is after 8 AM but before 1AM.
Or if the day of the week is Saturday or Sunday
These mark the cells so that they are calculated differently then the "day time" teachers.
"""
def create_time_range(start,end):
    lister=[start,end]
    time_range=""
    timeStart = '8:00PM'
    timeEnd = '1:00AM'
    timeEnd = datetime.strptime(timeEnd, "%I:%M%p")
    timeStart = datetime.strptime(timeStart, "%I:%M%p")
    for item in lister:
        hour=item[13:]
        date=item[9:12]
        string_time=datetime.strptime(hour,'%I:%M %p')
        if date =='Sat' or date=='Sun':
            hour=hour+'*'
        if timeStart<=string_time or string_time<=timeEnd:
            hour=hour+'*'
        time_range=time_range+hour+" - "
    return time_range[:-2]


start='04/25/18 Wed 4:12 PM'
end='04/25/18 Wed 5:54 PM'
expected_output='4:12 PM - 5:54 PM'
start1='04/25/18 Wed 5:54 PM'
end1='04/25/18 Wed 8:06 PM'
expected_output1='5:54 PM - 8:06 PM*'
start2='04/25/18 Wed 8:06 PM'
end2='04/25/18 Wed 11:18 PM'
expected_output2='8:06 PM* - 11:18PM*'
start3='04/25/18 Sat 4:12 PM'
end3='04/25/18 Sat 5:54 PM'
expected_output='4:12 PM* - 5:54 PM*'




