# Teaching Department Efficiency Report

The purpose of this script is to take an excel file with teacher ytd information
and output another excel sheet that is easy to read and includes an
efficiency score. First it makes a LeadBook, which is used by the Lead (a teaching department manager) to read information about all of the teachers on their team.
Then it creates individual workbooks for each teacher that only contains their personal information. 

The input excel file has a time stamp, share fairly number(Tabby), and each 
teachers Year to Date (YTD) total sessions taught at that time. The share fairly
number or Tabby, is the number of sessions a teacher should be teaching at that
time. I have provided the input files for you. You can use either 
FullTime_Team_Source.xlsx or PartTime_Team_Source.xlsx. Full time teaching teams do 
not work night shifts, so it will only return day data. 

Something to Note:    
Students Taught  ÷ Share Fairly Number = Efficiency Score

For example if Share Fairly number is 4, and a teacher took 2 students, then that teacher
has an efficiency score of .5

# Libraries
[openpyxl](https://openpyxl.readthedocs.io/en/stable/) is the main library used for this script. 

` pip install openpyxl `

# How to Run It
After libraries are installed, open up callmodule.py and edit the information
under user variables. Then in the command prompt run callmodule.py. Make sure you have
either FullTime_Team_Source.xlsx or PartTime_Team_Source.xlsx in the same folder as callmodule.

# How it Works

Here are the main modules and a short description of each. They run in the
following order. Some of these modules have some formatting modules attached to them that are not listed here.

**time_difference**: Goes through the YTD information and finds the difference between the year to date numbers.
 So if at 4:00 PM Stacey had a YTD of 2,000 and then at 4:06 Stacey hd a YTD to 2,004, that means between 4:00 and 4:06, Stacey took 4 students. 
 It places all of this information under a new excel tab called "Raw Changes"

**split_days**: Takes all of the data from Raw Changes and split it up into multiple worksheets, each worksheet for a different day. So all of data from 04-02-18, will go under a tab called 04-02 Mon.

**create_tables**: Goes into each invidiual day sheet and creates daily tables. There is a daily table for each teacher for each day.

**mark_blocks**: So a "block" of teaching starts when a teacher has been teaching for at least 12 minutes, and ends when they have stopped teaching for at least 12 minutes. The script goes through and bolds all of the numbers that are in an active block. It then passes that block over to calculate block_escore. Blocks are conditionally formatted for easy reading, but those colors are not actually used in the efficiency score at all.

**calculate_block_escore**: Takes an active block and find the average students taken every 6 minutes during that block, the average tabby every 6 minutes during the block, and divides them to get the block efficiency score. It then pastes that information in the tables created during the create_tables module. Then each block is conditionally formatted in the tables, using the condition list. 

**calculate_daily_escore**: Goes through the table and averages out the escore for all of the blocks during the day, prodcing a daily escore. Night time shifts (which are any shift that touches 
8PM) are calculated seperatly because they have a different "good score" (it's harder for them to reach higher numbers since the demand is lower at night). It then saves all of these daily and nightly escores into a dictionary. That final dictionary is passed over to efficiency_score_summary

**efficiency_score_summary**: Creates a new tab called "Summary" and is placed as the first worksheet. The Summary tab is suppose to be the summary of all of the data. 
It creates tables for whatever data is available. So it will only createa day summary table if there are only day shift, or a day and night table if there are both types of shifts. The table has dates as it's y axis and the teacher names as it's x axis. It then uses those axis, to find the relevant information in the dictionary passed over from calculate_daily_escore, and pastes them in.

**teacherbooks**: Creates the personal teacher books. It goes through each teacher in the summary page and saves it as teachername. It then uses that teachname and goes through each day of week worksheet and then find their information from each page and copie and pastes it into a newbook, with that teachers name as a title. 



# Known Issues
The FAQ image for the individual teacher books needs to be updated. It refrences 5 minute intervals instead of 6 minute ones.

<<<<<<< HEAD

# FAQ Page
![FAQ](https://raw.githubusercontent.com/civetta/Efficiency-Report/master/faq.png?token=AEV7slistXziNgm7lU3kQPBQw_O6Ww22ks5bMOXGwA%3D%3D)
=======
# FAQ Page
![FAQ](https://raw.githubusercontent.com/civetta/Efficiency-Report/master/faq.png?token=AEV7slistXziNgm7lU3kQPBQw_O6Ww22ks5bMOXGwA%3D%3D)
>>>>>>> 1123ca88d24ddfe49d265f74f7690606630f5527
