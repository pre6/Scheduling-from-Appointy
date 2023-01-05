from openpyxl import Workbook
from openpyxl import load_workbook
#from create_list import create_list
import datetime

def get_cell(lst,day_num):

    Alpha = '@ABCDEFGHIJKLMNOPQRSTUVWXYZabcedefghijklmnopqrstuvwxyz'

    if day_num == 0:
        initial_time = datetime.time(hour = 10)
    else:
        initial_time = datetime.time(hour = 15, minute = 30)

    initial_column = 'G'
    initial_row = (28 * (4-day_num)+6)

    used_col = []
    a = datetime.time(hour = 15, minute = 30)
    b = datetime.time(hour = 15, minute = 45)
    c = datetime.time(hour = 16)
    d = datetime.time(hour = 10)
    e = datetime.time(hour = 10, minute = 15)
    f = datetime.time(hour = 10, minute = 30)

    time_list = [a,b,c,d,e,f]

    for i in range(25):

        for column_dict in lst:

            start_time = list(column_dict.keys())[0]
            if start_time.time() == initial_time:

                if initial_time in time_list:
                    used_col.append(initial_column)

                    student = column_dict[start_time]
                    student.append(initial_column + str(initial_row))

                    num = Alpha.index(initial_column) + 3
                    initial_column = Alpha[num]

                else:
                    student = column_dict[start_time]
                    student.append(initial_column+initial_row)
                    used_col.append(initial_column)

        initial_row = ((28 * (4-day_num)+6) + i + 1)
        rand_day = datetime.date(2021, 1, 4)
        delta = datetime.timedelta(minutes = 15)
        new_date_with_time = datetime.datetime.combine(rand_day,initial_time)

        initial_time = (new_date_with_time + delta).time()


        if initial_time in time_list:
            num = Alpha.index('G') + i + 1
            initial_column = Alpha[num]

        else:
            for i in used_col:
                if i in Alpha:
                    Alpha.replace(i,'')
            initial_column = Alpha[0]
