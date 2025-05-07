from openpyxl import Workbook
from openpyxl import load_workbook
#from datetime import datetime, time, timedelta
from create_list import create_list
import datetime

x = create_list()
def split_list(lst):
    sat = []
    thurs = []
    wed = []
    tues = []
    mon = []

    day = []
    day.append(sat)
    day.append(thurs)
    day.append(wed)
    day.append(tues)
    day.append(mon)
    i = 0
    while len(lst) > 0:
        delete = lst.copy()
        date = list(lst[0].keys())[0].date()
        for item in lst:
            new = list(item.keys())[0].date()
            if new == date:
                day[i].append(item)
                delete.remove(item)
        lst = delete

        i += 1

    return day

y = split_list(x[1])


def get_cells_ol(lst,day_num):
    Alpha = 'CDEF'

    if day_num == 0:
        initial_time = datetime.time(hour = 10)
    else:
        initial_time = datetime.time(hour = 15, minute = 30)
    initial_column = 'C'
    initial_row = (28 * (4-day_num)+6)

    for i in range(22):

        for column_dict in lst:

            start_time = list(column_dict.keys())[0]

            if start_time.time() == initial_time:
                student = column_dict[start_time]
                student.append(initial_column)
                student.append(initial_row)
                num = Alpha.index(initial_column) + 1
                initial_column = Alpha[num]

        initial_row = ((28 * (4-day_num)+6) + i + 1)
        rand_day = datetime.date(2021, 1, 4)
        delta = datetime.timedelta(minutes = 15)
        new_date_with_time = datetime.datetime.combine(rand_day,initial_time)

        initial_time = (new_date_with_time + delta).time()
    return lst
'''
cells = []
i = 0
for item in y:
    f = get_cells_ol(item,i)
    cells.append(f)
    i += 1

for thing in cells:
    print(thing)
    print()

'''




def get_cell(lst,day_num):

    Alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcedefghijklmnopqrstuvwxyz'

    if day_num == 0:
        initial_time = datetime.time(hour = 10)
    else:
        initial_time = datetime.time(hour = 15, minute = 30)
    initial_column = 'G'
    initial_row = (28 * (4-day_num)+6)

    used_col = ['A','B','C','D','E','F']
    a = datetime.time(hour = 15, minute = 30)
    b = datetime.time(hour = 15, minute = 45)
    c = datetime.time(hour = 16)
    d = datetime.time(hour = 10)
    e = datetime.time(hour = 10, minute = 15)
    f = datetime.time(hour = 10, minute = 30)

    time_list = [a,b,c,d,e,f]

    for i in range(17):

        for column_dict in lst:

            start_time = list(column_dict.keys())[0]
            if start_time.time() == initial_time:

                if initial_time in time_list:
                    used_col.append(initial_column)

                    student = column_dict[start_time]
                    student.append(initial_column)
                    student.append(initial_row)

                    num = Alpha.index(initial_column) + 3
                    initial_column = Alpha[num]

                else:
                    used_col.append(initial_column)
                    student = column_dict[start_time]
                    student.append(initial_column)
                    student.append(initial_row)
                    num = Alpha.index(initial_column) + 1
                    initial_column = Alpha[num]


        initial_row = ((28 * (4-day_num)+6) + i + 1)
        rand_day = datetime.date(2021, 1, 4)
        delta = datetime.timedelta(minutes = 15)
        new_date_with_time = datetime.datetime.combine(rand_day,initial_time)

        initial_time = (new_date_with_time + delta).time()


        if initial_time in time_list:
            num = Alpha.index('G') + i + 1
            initial_column = Alpha[num]

        else:
            for n in used_col:
                if n in Alpha:
                    Alpha = Alpha.replace(n,'')
            initial_column = Alpha[0]

    return lst

'''
cells = []
i = 0
for item in y:
    f = get_cell(item,i)
    cells.append(f)
    i += 1

for thing in cells:
    print(thing)
    print()

'''
