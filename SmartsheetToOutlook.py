from email.errors import StartBoundaryNotFoundDefect
from icalendar import Calendar, vCalAddress, Event, vText
import smartsheet
import json
import pytz
from datetime import datetime
import os

####################################
# Fetch Smartsheet object from API #
####################################
SMARTSHEET_ACCESS_TOKEN = "cQRQeMelih0hMj0DHzx60l6Ggu5xcHvJvYMGU"
sheet_id = 1068857952626564

smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
sheet = smartsheet_client.Sheets.get_sheet(sheet_id) # JSON object

sheet = json.loads(str(sheet)) # Convert JSON object to Python object

#sheet_str = json.dumps(sheet, indent=4) # Get printable string from Python object
#print(sheet_str[:30000])

#################################################
# Create calendar event for each Smartsheet row #
#################################################
cal = Calendar()

for meeting in sheet['rows'][6:11]:
    event = Event()

    startTime = meeting['cells'][3]['value']
    endTime = meeting['cells'][4]['value']

    startHour = int(startTime/100)
    startMinute = int(startTime%100)
    endHour = int(endTime/100)
    endMinute = int(endTime%100)

    print('from',startHour,":",startMinute,'to',endHour,":",endMinute)

    print(meeting['cells'][13]['value'])
    print(meeting['cells'][14]['value'])

    startDate = meeting['cells'][13]['value'].split('-')
    endDate = meeting['cells'][14]['value'].split('-')

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