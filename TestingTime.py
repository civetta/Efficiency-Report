from datetime import datetime
"""
Input: Two string of time. Example: 03/21/18 Thu 8:27 AM and 03/21/18 Thu 9:48 AM.
Output: A string with just the hours (example: 8:27 AM - 9:48 AM").
Times is marked with an asteriks if one of the times is after 8 AM but before 1AM.
Or if the day of the week is Saturday or Sunday
These mark the cells so that they are calculated differently then the "day time" teachers later on.
"""
def create_time_range(start,end):
    lister=[start,end]
    time_range=""
    night_shift_indicator_start = datetime.strptime('8:00PM', "%I:%M%p").time()
    night_shift_indicator_end = datetime.strptime('2:00AM', "%I:%M%p").time()
    print night_shift_indicator_start
    for item in lister:
        hour=item[13:]
        date=item[9:12]
        string_time=datetime.strptime(hour,'%I:%M %p').time()
        if date =='Sat' or date=='Sun' or night_shift_indicator_start <= string_time or string_time<=night_shift_indicator_end:
            hour=hour+'*'
        time_range=time_range+hour+" - "
    return time_range[:-3]

start='04/25/18 Wed 4:12 PM'
end='04/25/18 Wed 5:54 PM'
expected_output='4:12 PM - 5:54 PM'
start1='04/25/18 Wed 5:54 PM'
end1='04/25/18 Wed 8:06 PM'
expected_output1='5:54 PM - 8:06 PM*'
start2='04/25/18 Wed 8:06 PM'
end2='04/25/18 Wed 11:18 PM'
expected_output2='8:06 PM* - 11:18 PM*'
start3='04/25/18 Sat 4:12 PM'
end3='04/25/18 Sat 5:54 PM'
expected_output3='4:12 PM* - 5:54 PM*'
start4='04/25/18 Mon 11:12 PM'
end4='04/25/18 Tue 12:54 AM'
expected_output4='11:12 PM* - 12:54 AM*'

import unittest
class MyTest(unittest.TestCase):
    def test(self):
        self.assertEqual(create_time_range(start,end), expected_output)
    def test1(self):
        self.assertEqual(create_time_range(start1,end1), expected_output1)
    def test2(self):
        self.assertEqual(create_time_range(start2,end2), expected_output2)
    def test3(self):
        self.assertEqual(create_time_range(start3,end3), expected_output3)
    def test4(self):
        self.assertEqual(create_time_range(start4,end4), expected_output4)

if __name__ == '__main__':
    unittest.main()



