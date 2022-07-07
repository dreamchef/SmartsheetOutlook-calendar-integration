import smartsheet
import json

PARTICIPANTS_COL = 16
FLAG_COL = 4
SD_COL = 11
ED_COL = 12
ST_COL = 2
ET_COL = 3

f = open('./SmartsheetAccessToken.txt')
SMARTSHEET_ACCESS_TOKEN = f.read()
f.close()
main_sheet_id = 1068857952626564

def findCollisions():
    smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
    sheet = smartsheet_client.Sheets.get_sheet(main_sheet_id) # JSON object
    main_sheet = json.loads(str(sheet)) # Convert JSON object to Python object

    row_map = {}
    i = 0

    for rows in sheet.rows:
        row_map[i] = rows.id
        i = i + 1
        
    column_map = {}
    for column in sheet.columns:
        column_map[column.title] = column.id

    rows_map = []

    conflict = False

    z = 0
    for meeting in main_sheet['rows']:
        for name in meeting['cells'][PARTICIPANTS_COL]['value'].split(", "):

            for rowToTestIndex in range(len(main_sheet['rows'])):

                rowToTest = main_sheet['rows'][rowToTestIndex] 

                if name in rowToTest['cells'][PARTICIPANTS_COL]['value'] and z is not rowToTestIndex:

                    print('found participant conflict')

                    if meeting['cells'][SD_COL]['value'] == rowToTest['cells'][SD_COL]['value']:

                        print('...found day conflict')

                        start = meeting['cells'][ST_COL]['value']
                        end = meeting['cells'][ET_COL]['value']

                        print(rowToTest['cells'])

                        compareStart = rowToTest['cells'][ST_COL]['value']
                        compareEnd = rowToTest['cells'][ET_COL]['value']

                        print(compareStart,start,compareEnd,'|||',compareStart,end,compareEnd)

                        if (start > compareStart and start < compareEnd) or (end > compareStart and end < compareEnd) or start == compareStart or end == compareEnd:

                            print('......found time conflict')

                            new_row = smartsheet.models.Row()
                            new_row.id = row_map[z]
                            new_row.cells.append({
                                'column_id': column_map['Collisions'],
                                'value': 1,
                                'strict': False
                            })

                            rows_map.append(new_row)

                            conflict = True

                if(conflict == True): break    
            if(conflict == True): break 
        if(conflict == False):
            new_row = smartsheet.models.Row()
            new_row.id = row_map[z]
            new_row.cells.append({
                'column_id': column_map['Collisions'],
                'value': 0,
                'strict': False
            })
            rows_map.append(new_row)
        z = z + 1
        conflict = False

    smartsheet_client.Sheets.update_rows(
        main_sheet_id,
        rows_map
    )

findCollisions()
