import xlrd
import os
import requests
dir_path = os.path.dirname(os.path.realpath(__file__))


def receive_file(url, username):
    if not os.path.isfile(dir_path + '/spreadsheets/' + username + 'schedule.xlsx') and url[-5:] == '.xlsx':
        download_file(url, username)
    else:
        return True


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

            row = None
            col = day + 1

            if semester == 1:
                if hour == 9:
                    if minute == 0:
                        row = 1
                    elif minute == 30:
                        row = 2
                elif hour == 10:
                    if minute == 0:
                        row = 3
                    elif minute == 30:
                        row = 4
                elif hour == 11:
                    if minute == 0:
                        row = 5
                    elif minute == 30:
                        row = 6
                elif hour == 12:
                    if minute == 0:
                        row = 7
                    elif minute == 30:
                        row = 8
                elif hour == 13:
                    if minute == 0:
                        row = 9
                    elif minute == 30:
                        row = 10
                elif hour == 14:
                    if minute == 0:
                        row = 11
                    elif minute == 30:
                        row = 12
                elif hour == 15:
                    if minute == 0:
                        row = 13
                    elif minute == 30:
                        row = 14
                elif hour == 16:
                    if minute == 0:
                        row = 15
                    elif minute == 30:
                        row = 16
                elif hour == 17:
                    if minute == 0:
                        row = 17
                    elif minute == 30:
                        row = 18
                elif hour == 18:
                    if minute == 0:
                        row = 19
                    elif minute == 30:
                        row = 20
                elif hour == 19:
                    if minute == 0:
                        row = 21
                    elif minute == 30:
                        row = 22
                elif hour == 20:
                    if minute == 0:
                        row = 23
                    elif minute == 30:
                        row = 24
                elif hour == 21:
                    if minute == 0:
                        row = 25
                    elif minute == 30:
                        row = 26
            elif semester == 2:
                if hour == 9:
                    if minute == 0:
                        row = 29
                    elif minute == 30:
                        row = 30
                elif hour == 10:
                    if minute == 0:
                        row = 31
                    elif minute == 30:
                        row = 32
                elif hour == 11:
                    if minute == 0:
                        row = 33
                    elif minute == 30:
                        row = 34
                elif hour == 12:
                    if minute == 0:
                        row = 35
                    elif minute == 30:
                        row = 36
                elif hour == 13:
                    if minute == 0:
                        row = 37
                    elif minute == 30:
                        row = 38
                elif hour == 14:
                    if minute == 0:
                        row = 39
                    elif minute == 30:
                        row = 40
                elif hour == 15:
                    if minute == 0:
                        row = 41
                    elif minute == 30:
                        row = 42
                elif hour == 16:
                    if minute == 0:
                        row = 43
                    elif minute == 30:
                        row = 44
                elif hour == 17:
                    if minute == 0:
                        row = 45
                    elif minute == 30:
                        row = 46
                elif hour == 18:
                    if minute == 0:
                        row = 47
                    elif minute == 30:
                        row = 48
                elif hour == 19:
                    if minute == 0:
                        row = 49
                    elif minute == 30:
                        row = 50
                elif hour == 20:
                    if minute == 0:
                        row = 51
                    elif minute == 30:
                        row = 52
                elif hour == 21:
                    if minute == 0:
                        row = 53
                    elif minute == 30:
                        row = 54
            elif semester == 0:
                if hour == 9:
                    if minute == 0:
                        row = 57
                    elif minute == 30:
                        row = 58
                elif hour == 10:
                    if minute == 0:
                        row = 59
                    elif minute == 30:
                        row = 60
                elif hour == 11:
                    if minute == 0:
                        row = 61
                    elif minute == 30:
                        row = 62
                elif hour == 12:
                    if minute == 0:
                        row = 63
                    elif minute == 30:
                        row = 64
                elif hour == 13:
                    if minute == 0:
                        row = 65
                    elif minute == 30:
                        row = 66
                elif hour == 14:
                    if minute == 0:
                        row = 67
                    elif minute == 30:
                        row = 68
                elif hour == 15:
                    if minute == 0:
                        row = 69
                    elif minute == 30:
                        row = 70
                elif hour == 16:
                    if minute == 0:
                        row = 71
                    elif minute == 30:
                        row = 72
                elif hour == 17:
                    if minute == 0:
                        row = 73
                    elif minute == 30:
                        row = 74
                elif hour == 18:
                    if minute == 0:
                        row = 75
                    elif minute == 30:
                        row = 76
                elif hour == 19:
                    if minute == 0:
                        row = 77
                    elif minute == 30:
                        row = 78
                elif hour == 20:
                    if minute == 0:
                        row = 79
                    elif minute == 30:
                        row = 80
                elif hour == 21:
                    if minute == 0:
                        row = 81
                    elif minute == 30:
                        row = 82

            if str(worksheet.cell_value(row, col)).lower() != 'busy':
                user_id = '%s' % xlsxfile
                user_id = user_id[:-13]
                available.append(int(user_id))

    return available
