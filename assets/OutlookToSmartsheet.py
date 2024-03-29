import datetime as dt
from turtle import update
from types import NoneType
import pandas as pd
import win32com.client
from smartsheet import smartsheet
import collisions

#Individual Access Token...Can be created on SmartSheets: Profile->Integrations->API Access->Build Token
f = open('./SmartsheetAccessToken.txt')
SMARTSHEET_ACCESS_TOKEN = f.readline()[:-1]
START_DATE = f.readline()[:-1]
END_DATE = f.readline()
f.close()

def delete_existing_data(client, sheet, chunk_interval = 300):
    deleteRows = [row.id for row in sheet.rows]
    for x in range(0, len(deleteRows),chunk_interval):
        client.Sheets.delete_rows(sheet.id,deleteRows[x:x+chunk_interval])
    print("SUCCESS")

def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Folders
    for m in calendar:
        if "Meetings from Smartsheet" in m.Name :
            calendar = m.Items
            print(calendar)
            calendar.IncludeRecurrences = True
            calendar.Sort('[Start]')
            restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
            calendar = calendar.Restrict(restriction)
            return calendar
        else:
            print("No calendar found from smartsheets")
        print(m.Name)

#View for 3 week period to gain all of the meetings from outlook
begin = dt.datetime.strptime(START_DATE, '%m/%d/%Y')
end = dt.datetime.strptime(END_DATE, '%m/%d/%Y')

cal = get_calendar(begin, end)

#Task Type Dictionary for updating participants based on type
task = {
  "label": "Interns",
  "type": "shell",
  "command": "Assign task",
  "dependsOrder": "sequence",
  "dependsOn": ["Dynamics", "Supply Chain"]
}

task2 = {
  "label": "Dynamics",
  "type": "shell",
  "command": "Assign task",
  "dependsOrder": "sequence",
  "dependsOn": ["Interns", "Supply Chain"]
}

task3 = {
  "label": "Supply Chain",
  "type": "shell",
  "command": "Assign task",
  "dependsOrder": "sequence",
  "dependsOn": ["Dynamics", "Interns"]
 }


# Find from specific sheet in smartsheet settings
sheet_id = 1068857952626564

smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

rows = []

column_map = {}
for column in sheet.columns:
    column_map[column.title] = column.id

rows_array = []
if(cal is not None):
    for meeting in cal:
        calStartDate = meeting.start.strftime("%m/%d/%Y")
        calTemp = int(meeting.start.strftime("%H:%M:%S")[0:2]) % 12
        calStartTime = meeting.start.strftime("%H:%M:%S")
        calFinalTime = str(calTemp) + ':' + calStartTime[3:-3]
        calAdditionalPart = meeting.OptionalAttendees.replace(";",',')
        calEndTime = meeting.end.strftime("%H:%M:%S")
        calEndDate = meeting.end.strftime("%m/%d/%Y")

        match meeting.categories.split(", ")[0]:
            case "Orange Category":
                calTask = "Interns"
            case "Blue Category":
                calTask = "Dynamics"
            case "Purple Category":
                calTask = "Supply Chain"
            case _:
                calTask = "Add category to OutlookToSmartsheet Code"
                
        calSubject = meeting.subject
        calBody = meeting.body
        calDuration = str(meeting.duration) + 'm'

        row_a = smartsheet_client.models.Row()
        row_a.to_top = True

        row_a.cells.append({
            'column_id': column_map['Meeting Name'],
            'value': calSubject,
            'strict': False
        })

        row_a.cells.append({
            'column_id': column_map['Task Type'],
            'value': calTask,
            'strict': False
        })

        row_a.cells.append({
            'column_id': column_map['Military Start Time'],
            'formula': '=IF(VALUE(LEFT([Start Time]@row, FIND(":", [Start Time]@row) - 1)) <> 12, IF(CONTAINS("a", [Start Time]@row), VALUE(LEFT([Start Time]@row, FIND(":", [Start Time]@row) - 1)) * 100 + VALUE(MID([Start Time]@row, FIND(":", [Start Time]@row) + 1, 2)), VALUE(LEFT([Start Time]@row, FIND(":", [Start Time]@row) - 1)) * 100 + VALUE(MID([Start Time]@row, FIND(":", [Start Time]@row) + 1, 2)) + 1200), IF(CONTAINS("a", [Start Time]@row), VALUE(LEFT([Start Time]@row, FIND(":", [Start Time]@row) - 1)) * 100 + VALUE(MID([Start Time]@row, FIND(":", [Start Time]@row) + 1, 2)) + 1200, VALUE(LEFT([Start Time]@row, FIND(":", [Start Time]@row) - 1)) * 100 + VALUE(MID([Start Time]@row, FIND(":", [Start Time]@row) + 1, 2))))',
            'strict': False
        })

        row_a.cells.append({
            'column_id': column_map['Military End Time'],
            'formula': '=IF(VALUE(LEFT([End Time]@row, FIND(":", [End Time]@row) - 1)) <> 12, IF(CONTAINS("a", [End Time]@row), VALUE(LEFT([End Time]@row, FIND(":", [End Time]@row) - 1)) * 100 + VALUE(MID([End Time]@row, FIND(":", [End Time]@row) + 1, 2)), VALUE(LEFT([End Time]@row, FIND(":", [End Time]@row) - 1)) * 100 + VALUE(MID([End Time]@row, FIND(":", [End Time]@row) + 1, 2)) + 1200), IF(CONTAINS("a", [End Time]@row), VALUE(LEFT([End Time]@row, FIND(":", [End Time]@row) - 1)) * 100 + VALUE(MID([End Time]@row, FIND(":", [End Time]@row) + 1, 2)) + 1200, VALUE(LEFT([End Time]@row, FIND(":", [End Time]@row) - 1)) * 100 + VALUE(MID([End Time]@row, FIND(":", [End Time]@row) + 1, 2))))',
            'strict': False
        })

        row_a.cells.append({
            'column_id': column_map['Duration'],
            'value': calDuration,
            'strict': False
        })

        if int(calStartTime[0:2]) < 12:
            calFinalTime += 'am'
        else:
            calFinalTime += 'pm'

        row_a.cells.append({
            'column_id': column_map['Start Time'],
            'value': calFinalTime,
            'strict': False
        })

        row_a.cells.append({
            'column_id': column_map['Start Date'],
            'value': calStartDate,
            'strict': False
        })
        
        # row_a.cells.append({
        #     'column_id': column_map['Participants'],
        #     'formula': '=VLOOKUP([Task Type]@row, {Group Members}, 2, false)',
        #     'strict' : False
        # })

        calAdditionalPart = calAdditionalPart.split(', ')
        print(calAdditionalPart)
        row_a.cells.append({
            'column_id': column_map['Additional Attendees'],
            'object_value':{'objectType': 'MULTI_PICKLIST', 'values': calAdditionalPart},
            'strict': False
        })

        row_a.cells.append({
            'column_id': column_map['Comments'],
            'value': calBody,
            'strict': False
        })
        rows_array.append(row_a)

    delete_existing_data(smartsheet_client,sheet,300)
    updated_row = smartsheet_client.Sheets.add_rows(sheet_id,rows_array)
    collisions.findCollisions()
    print("Loaded Sheet: " + sheet.name)
else:
    print("Import a calendar from smartsheets using SmartsheetToOutlook.py")
# Check for any errors
smartsheet_client.errors_as_exceptions(True)
