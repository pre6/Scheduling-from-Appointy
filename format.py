from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
from openpyxl.styles import PatternFill


def delete_bad_columns():
    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    sheet = wb['appointmentsReport']
    sheet.delete_cols(3, 2)
    sheet.delete_cols(4, 4)
    sheet.delete_cols(5, 7)
    sheet.delete_cols(6, 5)
    sheet.delete_rows(1,1)
    wb.save('appointmentsReport.xlsx')

def fix_intake():

    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    ws = wb['appointmentsReport']

    maxr = ws.max_row
    maxc = ws.max_column

    for r in range(1,maxr+1):

        y = ws['E'+str(r)].value
        y = y.replace('Intake form:\n','')
        y = y.replace('Student Name:','')
        y = y.replace('Student #2: ,\n','')
        y = y.replace('Student #2:','')
        y = y.replace('Student #3:','')
        y = y.replace('Student #2 Name (if applicable): ,','')
        y = y.replace('Student #3 Name (if applicable): ','')
        y = y.replace(',\n','')
        y = y.replace('\n','')

        ws['E'+str(r)].value = y

    wb.save('appointmentsReport.xlsx')

def format_date_time():
    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    ws = wb['appointmentsReport']

    ws.insert_cols(3)

    maxr = ws.max_row

    for i in range(1,maxr+1):
        column = 'A'
        row  = str(i)
        date = ws[column + row].value.date()
        time = ws['B'+row].value
        new = datetime.datetime.combine(date,time)

        ws['C' + row].value = new

    ws.delete_cols(1, 2)
    wb.save('appointmentsReport.xlsx')


def remove_students():
    wb = load_workbook(filename = 'highschool.xlsx')
    ws = wb['Sheet1']

    maxr = ws.max_row

    highschool_students = []

    for i in range(1,maxr+1):
        highschool_students.append(ws['A'+str(i)].value)

    wb_appoint = load_workbook(filename = 'appointmentsReport.xlsx')
    ws_appoint = wb_appoint['appointmentsReport']

    maxrow = ws_appoint.max_row

    delete = []

    for j in range(1,maxrow+1):
        student = ws_appoint['D' + str(j)].value
        for thing in highschool_students:
            if thing == student:
                delete.append(j)
    i = 0
    for thing in delete:

        ws_appoint.delete_rows(thing - i)
        i += 1


    wb_appoint.save('appointmentsReport.xlsx')

def no_student_name():
    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    ws = wb['appointmentsReport']

    maxr = ws.max_row
    empty = '-'

    for i in range(1,maxr+1):
        student = ws['D' + str(i)].value

        if student == empty:
            parent = ws['B' + str(i)].value
            ws['D' + str(i)].value = parent

    wb.save('appointmentsReport.xlsx')

def home_online():
    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    ws_ic = wb.create_sheet("In_Center")
    ws_ol = wb.create_sheet("Online")
    ws = wb['appointmentsReport']

    maxr = ws.max_row
    maxc = ws.max_column
    i = 1
    y = 1

    for r in range(1,maxr+1):
        if ws['C'+str(r)].value == 'In-Centre Session - 1 Hour(1) ':
            ws_ic['A' + str(i)].value = ws['A' + str(r)].value
            ws_ic['B' + str(i)].value = ws['B' + str(r)].value
            ws_ic['C' + str(i)].value = ws['C' + str(r)].value
            ws_ic['D' + str(i)].value = ws['D' + str(r)].value
            ws_ic['E' + str(i)].value = ws['E' + str(r)].value
            i += 1

        else:
            ws_ol['A' + str(y)].value = ws['A' + str(r)].value
            ws_ol['B' + str(y)].value = ws['B' + str(r)].value
            ws_ol['C' + str(y)].value = ws['C' + str(r)].value
            ws_ol['D' + str(y)].value = ws['D' + str(r)].value
            ws_ol['E' + str(y)].value = ws['E' + str(r)].value
            y += 1

    wb.save('appointmentsReport.xlsx')

def create_list():
    wb = load_workbook(filename = 'appointmentsReport.xlsx')
    ws_ic = wb['In_Center']
    ws_ol = wb['Online']

    lst_ic = []
    lst_ol = []
    maxr = ws_ic.max_row
    maxr_ol = ws_ol.max_row



    for r in range(1,maxr+1):
        dict = {}
        date = ws_ic['A'+str(r)].value
        dict[date] = [ws_ic['D' + str(r)].value]
        delta = datetime.timedelta(hours = 1)

        for i in range(r+1,maxr+1):
            next_date = ws_ic['A'+str(i)].value
            if next_date == None:
                break
            elif next_date == date + delta:
                dict[next_date] = [ws_ic['D' + str(i)].value]
                ws_ic.delete_rows(i,1)
                date = next_date
        if dict != {None:[None]}:
            lst_ic.append(dict)



    for t in range(1,maxr_ol+1):
        dict_ol = {}
        date = ws_ol['A'+str(t)].value
        dict_ol[date] = [ws_ol['D' + str(t)].value]
        delta = datetime.timedelta(hours = 1)

        for j in range(t+1,maxr_ol+1):
            next_date = ws_ol['A'+str(j)].value
            if next_date == None:
                break
            elif next_date == date + delta:
                dict_ol[next_date] = [ws_ol['D' + str(j)].value]
                ws_ol.delete_rows(j,1)
                date = next_date
        if dict_ol != {None:[None]}:
            lst_ol.append(dict_ol)

    return [lst_ic,lst_ol]

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


def get_cell(lst,day_num):

    Alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI']

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
                    Alpha.remove(n)

            initial_column = Alpha[0]

    return lst


def fill_cell_ol(y):

    wb = load_workbook(filename = 'Template.xlsx')
    ws = wb['Week']

    fill = PatternFill(start_color='EA9999', fill_type='solid')

    for day in y:
        for dict in day:
            col = list(dict.values())[0][1]
            row = list(dict.values())[0][2]

            for item in dict:
                student = dict[item][0]
                ws[col + str(row)].value =  student
                ws[col + str(row)].fill = fill
                ws[col + str(row+1)].fill = fill
                ws[col + str(row+2)].fill = fill
                ws[col + str(row+3)].fill = fill
                row = row + 4

    wb.save('Template.xlsx')


def fill_cell(y):
    wb = load_workbook(filename = 'Template.xlsx')
    ws = wb['Week']

    first = PatternFill(start_color='FCE5CD', fill_type='solid')
    second = PatternFill(start_color='D9EAD3', fill_type='solid')
    third = PatternFill(start_color='9FC5E8', fill_type='solid')
    fourth = PatternFill(start_color='EAD1DC', fill_type='solid')
    fifth = PatternFill(start_color='B4A7D6', fill_type='solid')
    sixth = PatternFill(start_color='FFD966', fill_type='solid')
    seventh = PatternFill(start_color='FFD966', fill_type='solid')
    eighth = PatternFill(start_color='FFD966', fill_type='solid')
    ninth = PatternFill(start_color='FFD966', fill_type='solid')

    colours = [first,second, third, fourth, fifth,sixth,seventh,eighth,ninth]

    others = PatternFill(start_color='FFD966', fill_type='solid')

    other_times = []
    other_times.append(datetime.time(hour = 10))
    other_times.append(datetime.time(hour = 10,minute = 15))
    other_times.append(datetime.time(hour = 10,minute = 30))
    other_times.append(datetime.time(hour = 15,minute = 30))
    other_times.append(datetime.time(hour = 15, minute = 45))
    other_times.append(datetime.time(hour = 16))



    for day in y:
        i=0
        h= 0

        for dict in day:
            col = list(dict.values())[0][1]
            row = list(dict.values())[0][2]

            if list(dict.keys())[0].time() not in other_times:
                fill = others
            elif h != list(dict.keys())[0].time():
                i = 0
                fill = colours[i]
            else:
                fill = colours[i]

            for item in dict:
                student = dict[item][0]
                ws[col + str(row)].value =  student
                ws[col + str(row)].fill = fill
                ws[col + str(row+1)].fill = fill
                ws[col + str(row+2)].fill = fill
                ws[col + str(row+3)].fill = fill
                row = row + 4
            i+=1
            h = list(dict.keys())[0].time()
    wb.save('Template.xlsx')



delete_bad_columns()
fix_intake()
format_date_time()
remove_students()
no_student_name()
home_online()

list_ol_ic = create_list()

list_ic_days = split_list(list_ol_ic[0])
list_ol_days = split_list(list_ol_ic[1])

finsihed_list_ol = []
i = 0
for item in list_ol_days:
    f = get_cells_ol(item,i)
    finsihed_list_ol.append(f)
    i += 1

finished_list_ic = []
j = 0
for thing in list_ic_days:
    f = get_cell(thing,j)
    finished_list_ic.append(f)
    j += 1

fill_cell_ol(finsihed_list_ol)
fill_cell(finished_list_ic)
