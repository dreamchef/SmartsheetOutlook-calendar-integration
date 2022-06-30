import datetime as dt
import pandas as pd
import smartsheet


####################################
# Fetch Smartsheet object from API #
####################################
SMARTSHEET_ACCESS_TOKEN = "cQRQeMelih0hMj0DHzx60l6Ggu5xcHvJvYMGU"
sheet_id = 1068857952626564

smartsheet_client = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

print(sheet)
x = input('...enter to continue...')

