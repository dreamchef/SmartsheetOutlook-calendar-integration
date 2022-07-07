import smartsheet
import json


SMARTSHEET_ACCESS_TOKEN = "xPj6dbHDFbLBKuVqU86zBw4xI3lJPLgA3tTzT"
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
    
    z = 0
    for meeting in main_sheet['rows']:
        for name in meeting['cells'][16]['value'].split(", "):
            for rowToTest in main_sheet['rows']:
                if name in rowToTest['cells'][16]['value']:
                    new_row = smartsheet.models.Row()
                    new_row.id = row_map[z]
                    new_row.cells.append({
                        'column_id': column_map['Participant Collision'],
                        'value': 1,
                        'strict': False
                    })
                    rows_map.append(new_row)
        z = z + 1
    updated_row = smartsheet_client.Sheets.update_rows(
        main_sheet_id,
        rows_map
    )

findCollisions()
