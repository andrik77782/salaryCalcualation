import win32com.client as win32
from datetime import datetime, timedelta
from itertools import groupby
from funct import get_activeTasks, get_users, get_projects, departments


def fillDaysAndFormat(sheet, today):
    sheet.Columns("A").ColumnWidth = 11
    sheet.Rows("1").WrapText = True
    sheet.Columns("A:BB").ColumnWidth = 11
    sheet.Rows("1:1").Font.Bold = True
    sheet.Rows("1:1").RowHeight = 34
    sheet.Columns("B:BB").NumberFormatLocal = "0,00%"
    num = 3
    while num < 32:
        if today.weekday() < 5:
            sheet.Cells(num, "A").Value = str(today)
        today = today + timedelta(days=1)
        num += 1


def get_username(iterable):
    return iterable['user_to']['name']


def get_department(iterable):
    return str(usersNamesDepartments.get(iterable['user_to']['name'], 0))


def count_working_days(start_date, end_date):
    working_days = 0
    current_date = start_date
    
    while current_date <= end_date:
        if current_date.weekday() < 5:
            working_days += 1
        current_date += timedelta(days=1)
    
    return working_days


# створюємо excel файл
xlApp = win32.Dispatch('Excel.application')
xlApp.Visible = True
wb = xlApp.Workbooks.Add()
sheet = wb.Worksheets(1)
today = datetime.today().date()

# отримання списка співробітників для аналізу
users = get_users()['data']
users = list(filter(lambda user: user['department'] in departments, users))
usersNames = [user['name'] for user in users]
usersNamesDepartments = {}
for user in users:
    usersNamesDepartments[user['name']] = user['department']
    
#отримуємо список активних задач
projects = get_projects()['data']
allActiveTasks = []
for project in projects:
    tasks1 = get_activeTasks(project['page'])['data']
    for task1 in tasks1:
        try:
            if task1['child']:
                for task2 in task1['child']:
                    try:
                        if task2['child']:
                            for task3 in task2['child']:
                                allActiveTasks.append(task3)
                    except KeyError:
                        None
        except KeyError:
            None

# сортуємо по департментам активні задачі
usersTasks = []
allActiveTasks.sort(key=get_department)
for department, data in groupby(allActiveTasks, key=get_department):
    if department in usersNamesDepartments.values():
        allActiveTasksByDep = list(data)
        allActiveTasksByDep.sort(key=get_username)
        tasksByDep = []
        for user, data2 in groupby(allActiveTasksByDep, key=get_username):
            if usersNamesDepartments.get(user,0):
                tasksByDep.append([list(data2),user, usersNamesDepartments.get(user,0)])
        usersTasks.append([department, tasksByDep])

# 
tasksForDay = []
for depTasks in usersTasks:
    currentColumn = 2
    nextSheet = wb.Sheets.Add(After=sheet)
    sheet = nextSheet
    sheet.Name = depTasks[0]
    today = datetime.today().date()
    fillDaysAndFormat(sheet, today)  
    for userTasks in depTasks[1]:
        num, currentRow = 3, 1    
        today = datetime.today().date()    
        sheet.Cells(currentRow, currentColumn).Value = userTasks[1]
        currentRow += 1
        while num < 32:
            currentRow += 1    
            if today.weekday() < 5:             
                coefPerDay = 0
                for task in userTasks[0]:
                    try:
                        date_start = datetime.strptime(task['date_start'], "%Y-%m-%d").date()
                        date_end = datetime.strptime(task['date_end'], "%Y-%m-%d").date()
                        if date_start <= today <= date_end:
                            coefPerDay += task['max_time'] / (count_working_days(date_start, date_end) * 8)
                    except:
                        None
                sheet.Cells(currentRow, currentColumn).Value = coefPerDay
                if 1.2 > coefPerDay >= 1.1:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*192 + 256*192 + 246
                elif 1.3 > coefPerDay >= 1.2:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*165 + 256*165 + 246
                elif 1.4 > coefPerDay >= 1.3:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*120 + 256*120 + 237
                elif 1.5 > coefPerDay >= 1.4:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*98 + 256*98 + 234
                elif 1.6 > coefPerDay >= 1.5:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*75 + 256*75 + 231
                elif 1.7 > coefPerDay >= 1.6:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*53 + 256*53 + 227
                elif 1.8 > coefPerDay >= 1.7:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*31 + 256*31 + 224
                elif coefPerDay >= 1.8:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*28 + 256*28 + 208
                
                if 0.9 > coefPerDay >= 0.8:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*230 + 256*242 + 255
                elif 0.8 > coefPerDay >= 0.7:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*204 + 256*229 + 255
                elif 0.7 > coefPerDay >= 0.6:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*179 + 256*215 + 255
                elif 0.6 > coefPerDay >= 0.5:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*153 + 256*201 + 255
                elif 0.5 > coefPerDay >= 0.4:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*128 + 256*187 + 255
                elif 0.4 > coefPerDay >= 0.3:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*102 + 256*173 + 255
                elif 0.3 > coefPerDay >= 0.2:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*77 + 256*160 + 255
                elif 0.2 > coefPerDay:
                    sheet.Cells(currentRow, currentColumn).Interior.Color = 256**2*51 + 256*146 + 255

                
            today = today + timedelta(days=1)
            num += 1
        currentColumn += 1
