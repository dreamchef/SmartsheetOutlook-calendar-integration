from email.errors import StartBoundaryNotFoundDefect
from icalendar import Calendar, vCalAddress, Event, vText
import smartsheet
import json
from datetime import datetime
import os

SD_COL = 12
ED_COL = 13
ST_COL = 2
ET_COL = 3

START_ROW = 17
END_ROW = 19

NUM_GROUPS = 3

####################################
# Fetch Smartsheet object from API #
####################################
SMARTSHEET_ACCESS_TOKEN = "cQRQeMelih0hMj0DHzx60l6Ggu5xcHvJvYMGU"
main_sheet_id = 1068857952626564

group_sheet_id = 6814705466533764

smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
main_sheet = smartsheet_client.Sheets.get_sheet(main_sheet_id) # JSON object
group_sheet = smartsheet_client.Sheets.get_sheet(group_sheet_id)

main_sheet = json.loads(str(main_sheet)) # Convert JSON object to Python object
group_sheet = json.loads(str(group_sheet))

sheet_str = json.dumps(main_sheet, indent=4) # Get printable string from Python object
print(sheet_str[:30000])

#################################################
# Create calendar event for each Smartsheet row #
#################################################
cal = Calendar()

for meeting in main_sheet['rows'][START_ROW:END_ROW]:
    event = Event()

    categoryArray = []

    group = meeting['cells'][1]['value']

    groupRow = [rowIndex for rowIndex in range(NUM_GROUPS) if group_sheet['rows'][rowIndex] == group]

    categoryArray = [groupRow[0] + " Category"]

    meeting['cells'][1]['value']

    event.add('summary', meeting['cells'][0]['value'])
    event.add('description', group)
    event.add('categories',categoryArray)

    startTime = meeting['cells'][ST_COL]['value']
    endTime = meeting['cells'][ET_COL]['value']

    startHour = int((startTime/100)%24)
    startMinute = int(startTime%100)
    endHour = int(endTime/100)
    endMinute = int(endTime%100)

    #print('from',startHour,":",startMinute,'to',endHour,":",endMinute)

    #print(meeting['cells'][SD_COL]['value'])
    #print(meeting['cells'][ED_COL]['value'])

    startDate = meeting['cells'][SD_COL]['value'].split('-')
    endDate = meeting['cells'][ED_COL]['value'].split('-')

    startYear = int(startDate[0])
    startMonth = int(startDate[1])
    startDay = int(startDate[2][:2])
    endYear = int(endDate[0])
    endMonth = int(endDate[1])
    endDay = int(endDate[2][:2])

    event.add('dtstart', datetime(startYear, startMonth, startDay, startHour, startMinute, 0, 0))
    event.add('dtend', datetime(startYear, endMonth, endDay, startHour, endMinute, 0, 0))
    cal.add_component(event)

#####################################
# Write calendar events to ICS file #
#####################################
directory = "./"
print("ics file will be generated at ", directory)
f = open(os.path.join(directory, 'Meetings From Smartsheet.ics'), 'wb')
f.write(cal.to_ical())
f.close()