from openpyxl import Workbook
import os
import gspread
import time
import oauth2client   

#Completes all of the authorization for google sheets
def authorize():
    from oauth2client.service_account import ServiceAccountCredentials
    scope = "https://spreadsheets.google.com/feeds"
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:\\Users\\kheyden\\Documents\\Program\\archivekey.json', scope)
    gs = gspread.authorize(credentials)
    gsheet = gs.open("Test")
    wsheet = gsheet.worksheet("apr23")
    return wsheet


#The below three functions are utility functions that will be used repeatedily 
#########################################################################################
#Creates excel workbook
def create_workbooks():
    wb = Workbook()
    return wb

#Creates excel worksheet
def creat_worksheets(wb):
    ws=wb.create_sheet('Raw Pulls')
    return ws

#Saves the workbook    
"""def workbook_save(wb,ws):
        test_save_name = os.path.join('C:\\Users\\kheyden\\Documents\\Program\\2017\\WeeklySummary\\', 'Leadbook' + '.xlsx')
        wb.save(test_save_name)"""
#########################################################################################

        
#Takes what is inside of the google spreadsheet and pastes it inside of the excel sheet   
def copy_in_excel(wsheet,wb,ws,ranger,n):
    col_count = wsheet.col_count
    col = ranger[0]+1
    while col <= ranger[1]:
        excel_col=ws.max_column
        
        list_of_values = wsheet.col_values(col)
        for i in list_of_values:
            try:
                i=int(i)
            except:
                i=str(i)
        row_count = len(list_of_values)
        row = 1
        while row < row_count:     
            ws.cell(row = row, column = excel_col+n,value = list_of_values[row-1])
            
            row = row+1
        col = col+1
        n=1
    #workbook_save(wb,ws)


def find_team_range(wsheet,start,end):
    names=wsheet.row_values(1)
    start=names.index(start)
    end=names.index(end)
    return [start,end]
    
#Removes all of the data from the google spreadsheet, leaving only the top row with labels, and 3 empty rows
def wsheet_clean_up(wsheet):
    wsheet.resize(rows = 1,cols = wsheet.col_count)
    wsheet.add_rows(3)

#Calls all of the other functions above
def call_function(start,end):
    wsheet = authorize()
    wb=create_workbooks()
    ws=creat_worksheets(wb)
    #workbook_save(wb,ws)
    ranger=find_team_range(wsheet,start,end)
    copy_in_excel(wsheet,wb,ws,[0,2],0)
    copy_in_excel(wsheet,wb,ws,ranger,1)
    #workbook_save(wb,ws)
    #wsheet_clean_up(wsheet)
    return wb


    
