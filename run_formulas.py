import calendar
from datetime import datetime

import httplib2
from apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials

from formulas import get_use_left, get_statistics, get_my_statistics


TABLE_FORMULAS = [
    ("UseLeft!", get_use_left),
    ("Statistics!", get_statistics),
    ("Yura\'s Statistics!", get_my_statistics)
]

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
        majorDimension=dimension,  # 'ROWS',
        valueRenderOption=render_option,  # FORMULA
    ).execute()
    return values


def write_data(excel_range, dimension, values, render_option=None):
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "responseValueRenderOption": render_option,
            "data": [
                {
                    "range": excel_range,
                    "majorDimension": dimension,
                    "values": values,
                },
            ]
        }).execute()
    return values


def get_month_year_new_row(table_name):
    months = get_data(f'{table_name}A1:A', 'COLUMNS')
    excel_months_list = months['values'][0]

    new_row = len(excel_months_list) + 1

    year = datetime.now().year
    month = datetime.now().month
    month_name = calendar.month_name[datetime.now().month]
    return new_row, year, month, month_name


def update_table(table_name, formula_function):
    new_row, year, month, month_name = get_month_year_new_row(table_name)
    write_data(f'{table_name}A{new_row}', 'ROWS', [[month_name]])

    data = formula_function(year, month, new_row)
    data = [[d for d in data.values()]]
    formula_write_data = {
        'excel_range': f'{table_name}B{new_row}:{new_row}',
        'dimension': 'ROWS',
        'values': data,
        'render_option': 'FORMULA',
    }
    write_data(**formula_write_data)


def make_changes():
    for table_name, func in TABLE_FORMULAS:
        update_table(table_name, func)
