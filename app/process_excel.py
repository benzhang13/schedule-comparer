import xlrd
import os
import requests
from app.excel_cells import EXCEL_KEYS
dir_path = os.path.dirname(os.path.realpath(__file__))


def receive_file(url, username):
    if not os.path.exists(dir_path + '/spreadsheets'):
        os.mkdir(dir_path + '/spreadsheets')
    if os.path.isfile(dir_path + '/spreadsheets/' + username + 'schedule.xlsx'):
        return {'already_exists': True}
    if url[-5:] != '.xlsx':
        return {'not_xlsx': True}
    download_file(url, username)
    return {'submitted': True}


def download_file(url, username):
    r = requests.get(url)

    with open(dir_path + '/spreadsheets/' + username + 'schedule.xlsx', 'wb') as f:
        f.write(r.content)


def delete_file(username):
    if not os.path.isfile(dir_path + '/spreadsheets/' + username + 'schedule.xlsx'):
        return False
    else:
        os.remove(dir_path + '/spreadsheets/' + username + 'schedule.xlsx')
        return True


def read_all_files(hour, minute, day, semester):

    available = []

    for root, dirs, files in os.walk(dir_path + '/spreadsheets'):
        xlsxfiles = [_ for _ in files if _.endswith('.xlsx')]

        for xlsxfile in xlsxfiles:
            workbook = xlrd.open_workbook(os.path.join(root, xlsxfile))
            worksheet = workbook.sheet_by_index(0)

            if minute < 30:
                minute = 0
            else:
                minute = 1

            row = EXCEL_KEYS[semester][hour][minute]
            col = day + 1

            if str(worksheet.cell_value(row, col)).lower() != 'busy':
                user_id = '%s' % xlsxfile
                user_id = user_id[:-13]
                available.append(int(user_id))

    return available
