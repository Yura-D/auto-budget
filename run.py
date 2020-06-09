from pprint import pprint
import calendar
from datetime import datetime

import httplib2
from  apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials

from formulas import get_use_left


# Google Developer Console
CREDENTIALS_FILE = 'creds.json'
spreadsheet_id = '1u8iMJMs0o1Ak3dvpKSu9CTkjKVumYJ_toJvP0dsBdGw'

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = discovery.build('sheets', 'v4', http=httpAuth)


def get_data(excel_range, dimension, render_option=None):
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=excel_range,
        majorDimension=dimension, # 'ROWS',
        valueRenderOption=render_option, # FORMULA
    ).execute()
    return values


def write_data(excel_range, dimension, values, render_option=None):
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "responseValueRenderOption": render_option,
            "data": [
                {"range": excel_range,
                "majorDimension": dimension,
                "values": values},
            ]
        }).execute()
    return values


months = get_data('UseLeft!A1:A', 'COLUMNS')
excel_months_list = months['values'][0]

new_row = len(excel_months_list) + 1

last_month = excel_months_list[-1]
months_list = list(calendar.month_name)
year = datetime.now().year
month = datetime.now().month


month_name = calendar.month_name[datetime.now().month]

write_data(f'UseLeft!A{new_row}', 'ROWS', [[month_name]])

use_left_data = get_use_left(year, month, new_row)
use_left_data = [[data for data in use_left_data.values()]]
formula_write_data = {
    'excel_range': f'UseLeft!B{new_row}:{new_row}',
    'dimension': 'ROWS',
    'values': use_left_data,
    'render_option': 'FORMULA',
}
write_data(**formula_write_data)




"""


if last_month == months_list[12]:
    write_data(month_row, 'ROWS', [[datetime.now().year]])
    new_row += 1
    month_row = f'UseLeft!A{new_row}'
    write_data(month_row, 'ROWS', [[months_list[1]]])
    formula_write_data = {
        'excel_range': f'UseLeft!B{new_row}:{new_row}',
        'dimension': 'ROWS',
        'values': 'hello',
        'render_option': 'FORMULA',
    }
    write_data(**formula_write_data)



values = service.spreadsheets().values().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": "B3:C4",
             "majorDimension": "ROWS",
             "values": [["This is B3", "This is C3"], ["This is B4", "This is C4"]]},
            {"range": "D5:E6",
             "majorDimension": "COLUMNS",
             "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
	]
    }
).execute()
"""