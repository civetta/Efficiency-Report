import openpyxl
from openpyxl import load_workbook

#To Do - Do Not Paste Rows were there are 0,0,0,0,0,0 create start day switch. When turned on, copy everything.
def find_days(wb):

    search_day=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    ws = wb.get_sheet_by_name("Raw Changes")
    max_row=ws.max_row
    max_column=ws.max_column
    day_count=0
    rower=1
    a=2
    make_sheets(wb,ws,max_column,search_day[day_count])
    while a < max_row:
        while day_count<=6:
            day_sheet=wb.get_sheet_by_name(search_day[day_count])
            if ws.cell(row=a,column=1).value == None:
                break
            if search_day[day_count] in ws.cell(row=a,column=1).value:
                rower=rower+1

                for col in range(1,max_column+1):
                    valv=ws.cell(row=a, column=col).value
                    day_sheet.cell(row=rower, column=col, value=valv)
            else:
                day_count=day_count+1
                day=search_day[day_count]
                make_sheets(wb,ws,max_column,day)
                rower=1
            a=a+1
    return wb
        

def make_sheets(wb,ws,max_column,day):
    day_sheet = wb.create_sheet(day)
    for a in range(max_column):
        val=ws.cell(row=1,column=a+1).value
        day_sheet.cell(row=1,column=a+1,value=val)
        #wb.save('LeadBook.xlsx')
    
            
