import time
start = input("Start Date: (YYYY-MM-DD Formate):  ")
end = input("End Date: (YYYY-MM-DD Formate):  ")
Team = input("""Which Team? 1-Jeremy, 2-Caren,  3-All:  """)
Check = input("Are you connected to the database remotely via FortiClient? T or F: ")
check2 = input ("Do the dates you inputed above match what is in periscope? T or F: ")
boolean = input('Does the information you inputed, match what is currently in Periscope? T or F: ')
local = input('Are these reports to be put into official teacher one drive folders? If false, the reports will be saved locally to your computer and not put into one drive. T or F: ')

print "Collecting Data"
time.sleep(2)
print "Splitting Days"
time.sleep(1)
print "Calculating Blocks"
time.sleep(2)
print "Creating Summary Pages for Leadbook"
time.sleep(2)
print "Creating Teacherbooks"
time.sleep(4)
print "Efficiency Report Complete"