from ics import Calendar
from shutil import copyfile
from app.excel_cells import EXCEL_KEYS
import datetime
import openpyxl
import requests
import os
dir_path = os.path.dirname(os.path.realpath(__file__))


def receive_file(url, user_id, semester):
    if url[-4:] != '.ics':
        return {'not_ics': True}
    return write_to_xlsx(url, user_id, semester)


def write_to_xlsx(url, user_id, semester):
    user_id = str(user_id)
    if not os.path.isfile(dir_path + '/spreadsheets/' + user_id + 'schedule.xlsx'):
        copyfile(dir_path + '/generic_sheet/generic_sheet.xlsx',
                 dir_path + '/spreadsheets/' + user_id + 'schedule.xlsx')

    c = Calendar(requests.get(url).text)
    workbook = openpyxl.load_workbook(filename=dir_path+'/spreadsheets/'+user_id+'schedule.xlsx')
    worksheet = workbook.active
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    clear_alph = 'BCDEFGH'
    for key in EXCEL_KEYS[semester]:
        for row in EXCEL_KEYS[semester][key]:
            for col in clear_alph:
                xlformat = col + str(row+1)
                worksheet[xlformat] = ''

    for event in c.events:
        start = event.begin.datetime
        end = event.end.datetime

        while start < end:
            start_hour = start.hour
            start_minute = start.minute

            if start_minute < 30:
                start_minute = 0
            else:
                start_minute = 1

            col = start.weekday() + 1
            row = EXCEL_KEYS[semester][start_hour][start_minute] + 1
            xlformat = alphabet[col] + str(row)
            worksheet[xlformat] = 'busy'

            start = start + datetime.timedelta(minutes=30)
    workbook.save(dir_path + '/spreadsheets/' + user_id + 'schedule.xlsx')

    return {'submitted': True}


