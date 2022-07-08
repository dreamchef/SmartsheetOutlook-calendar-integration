from email.errors import StartBoundaryNotFoundDefect
from icalendar import Calendar, vCalAddress, Event, vText
import smartsheet
import json
from datetime import datetime
import os

# Column numbers of dates and times
SD_COL = 11
ED_COL = 12
ST_COL = 2
ET_COL = 3

####################################
# Fetch Smartsheet object from API #
####################################
f = open('./SmartsheetAccessToken.txt')
SMARTSHEET_ACCESS_TOKEN = f.read()
f.close()
main_sheet_id = 1068857952626564
group_sheet_id = 6814705466533764

smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
main_sheet = smartsheet_client.Sheets.get_sheet(main_sheet_id) # JSON object
group_sheet = smartsheet_client.Sheets.get_sheet(group_sheet_id)

main_sheet = json.loads(str(main_sheet)) # Convert JSON object to Python object
group_sheet = json.loads(str(group_sheet))

#sheet_str = json.dumps(group_sheet, indent=4) # Get printable string from Python object
#print(sheet_str[:30000])

#################################################
# Create calendar event for each Smartsheet row #
#################################################
cal = Calendar()

for meeting in main_sheet['rows']:
    event = Event()

    print(meeting['cells'])

    ### Assign correct color and name categories ###
    group = meeting['cells'][1]['value']
    numGroups = len(group_sheet['rows'])
    groupRow = [rowIndex for rowIndex in range(numGroups) if group_sheet['rows'][rowIndex]['cells'][0]['value'] == group]
    categoryArray = [group_sheet['rows'][groupRow[0]]['cells'][2]['value'] + " Category"]
    categoryArray.append(group)
    event.add('categories',categoryArray)

    ### Assign meeting title ###
    event.add('summary', meeting['cells'][0]['value'])

    ### Assign description from comments in Smartsheet ###
    if 'value' in meeting['cells'][18]:
        event.add('description', meeting['cells'][18]['value'])

    ### Assign meeting start/end date/time ###
    startTime = meeting['cells'][ST_COL]['value']
    endTime = meeting['cells'][ET_COL]['value']

    startHour = int((startTime/100)%24)
    startMinute = int(startTime%100)
    endHour = int(endTime/100)
    endMinute = int(endTime%100)

    startDate = meeting['cells'][SD_COL]['value'].split('-')
    endDate = meeting['cells'][ED_COL]['value'].split('-')

    startYear = int(startDate[0])
    startMonth = int(startDate[1])
    startDay = int(startDate[2][:2])
    endYear = int(endDate[0])
    endMonth = int(endDate[1])
    endDay = int(endDate[2][:2])

    event.add('dtstart', datetime(startYear, startMonth, startDay, startHour, startMinute, 0, 0))
    event.add('dtend', datetime(startYear, endMonth, endDay, endHour, endMinute, 0, 0))
    ### ---------------------------------- ###

    cal.add_component(event)
    cal.add('x-wr-calname','Meetings from Smartsheet')

#####################################
# Write calendar events to ICS file #
#####################################
directory = "./"
print("ics file will be generated at ", directory)
f = open(os.path.join(directory, 'Meetings From Smartsheet.ics'), 'wb')
f.write(cal.to_ical())
f.close()