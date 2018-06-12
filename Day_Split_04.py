def split_sheet_by_days(wb, skip_days):
    """Find start of day, and end of day, and then copies everything in
    between into a new sheet. If end of day returns none, it breaks"""
    raw_change_ws = wb.get_sheet_by_name("Raw Changes")
    max_row = raw_change_ws.max_row
    max_column = raw_change_ws.max_column+1
    end_of_day_row = 0
    day_count = 0
    #Keeps going for as long as current row is not equal or greater than max_row.
    while True:
        start_row = end_of_day_row+1
        if start_row >= max_row:
            break
        current_time_cell = raw_change_ws.cell(row=start_row, column=1).value
        if any(skip in current_time_cell for skip in skip_days):
            start_row = find_next_day(raw_change_ws, start_row, max_row)
        start_row = find_start(raw_change_ws, start_row, max_column, max_row)
        end_of_day_row = find_end(raw_change_ws, start_row, max_column, max_row)
        if start_row == end_of_day_row:
            #If current row is the same as end of day row, start the loop over 
            #again, meaning it looks for the following day because the day
            #had no teaching in it. It was blank
            pass
        if end_of_day_row is None:
            end_of_day_row = max_row
        if start_row >= max_row:
            break
        current_day = find_current_day(raw_change_ws, start_row)
        print 'Start Row: '+str(start_row)
        print 'End Row: '+str(end_of_day_row)
        print 'Day: '+str(current_day)
        print '\n'
        make_sheets(wb, start_row, end_of_day_row, raw_change_ws, max_column, current_day,day_count)
        day_count = day_count+1
    return wb


def find_next_day(raw_change_ws, start_row, max_row):
    """If skip day is used, then this function finds the next day"""
    current_row = start_row
    while current_row < max_row:
        current_day_time = raw_change_ws.cell(row=current_row, column=1).value
        if "12:54 AM" in current_day_time:
            return current_row+1
        else:
            current_row = current_row+1


def find_current_day(raw_change_ws, start_row):
    """Splices down to day of week. Used to name worksheet"""
    current_day_time = raw_change_ws.cell(row=start_row, column=1).value
    current_day_time = current_day_time[current_day_time.index(" ")+1:]
    current_day = current_day_time[:current_day_time.index(" ")]
    return current_day


def find_start(ws, start_row, max_column, max_row):
    """Find the first row with at least 3 consecutive values above 0"""
    current_row = start_row
    while current_row < max_row:
        for col in range(3, max_column+1):
            val1 = ws.cell(row=current_row, column=col).value
            try:  
                if int(val1) > 0:
                    val2 = int(ws.cell(row=current_row+1, column=col).value)
                    val3 = int(ws.cell(row=current_row+2, column=col).value)
                    if val2 > 0 or val3 > 0:
                        return current_row
            except:
                continue
        current_row = current_row + 1
    return current_row


def find_end(raw_change_ws, start_row, max_column, max_row):
    """Finds next instance of 12:45AM, 
    used to define the end row, or end of day."""    
    current_row = start_row
    while current_row < max_row:
        current_day_time = raw_change_ws.cell(row=current_row, column=1).value
        if "12:54 AM" in current_day_time:
            return current_row+1
        else:
            current_row = current_row+1


def make_sheets(wb, start_row, end_of_day_row, raw_change_ws, max_column, day, day_count):
    """Creates Worksheet with Day of Week-Day of Month title syntax. Copies and
     Pastes from Raw Changes using Start Row and End Row as ranges. Day count
     is used to help with sheet order"""
    name = raw_change_ws.cell(row=start_row, column=1).value
    name = name.replace("/", "-")
    current_sheet = wb.create_sheet(day+" "+name[:5],day_count)
    for column in range(1, max_column):
        name_cell = raw_change_ws.cell(row=1, column=column).value
        current_sheet.cell(row=1, column=column+5, value=name_cell)
        rower = 2
        for row in range(start_row, end_of_day_row):
            current_cell = raw_change_ws.cell(row=row, column=column).value
            current_sheet.cell(row=rower, column=column+5, value=current_cell)
            rower = rower+1
    return current_sheet
