import win32com.client as win32
import os, datetime
from itertools import groupby
from funct import get_users, get_costs_by_dates, get_costs, toDecimal, get_tasks, departments


def get_projnames(iterable):
    return [iterable['task']['project']['name'],
            iterable['task']['project']['page']]


def get_taskid(iterable):
    return iterable['task']['id']


def get_taskname(iterable):
    return iterable[0][0]


def get_taskname2(iterable):
    return iterable[0][1]


def get_taskname3(iterable):
    return iterable[0][2]


def get_department(iterable):
    return iterable['department'] 


def createExcel(firstWord, path, date, monthName):

    # створюємо excel файл та зберігаємо його
    wb = xlApp.Workbooks.Add()
    xlApp.Visible = True
    wb.SaveAs(os.path.join(path, '{}{}.xlsx'.format(firstWord, monthName)))
    sheet = wb.Worksheets(1)
    sheet.Columns("A:B").ColumnWidth = 2.4
    sheet.Columns("D:F").ColumnWidth = 10
    sheet.Columns("C").ColumnWidth = 40
    sheet.Range("A1 : O1").Font.Bold = True

    # заповнюємо шапку в ексель
    sheet.Cells(1, "A").Value = date
    sheet.Cells(1, "D").Value = "% від ставки"
    sheet.Cells(1, "E").Value = "Ставка"
    sheet.Cells(1, "F").Value = "До виплати"

    return sheet, wb


xlApp = win32.Dispatch('Excel.application')
xlApp.Visible = False

# основні дані
rootPath = "C:\\WS, bills\\temp"
nameMonth = " січень 2024"
dateFrom = datetime.datetime(2024, 1, 1)
dateTo = datetime.datetime(2024, 1, 31)
workingDays = 22
days = (dateTo - dateFrom).days + 1


# отримання списка співробітників та їх витрат за певний період
users = get_users()['data']
df = "'{}-{:02}-{:02}'".format(dateFrom.year, dateFrom.month, dateFrom.day)
dt = "'{}-{:02}-{:02}'".format(dateTo.year, dateTo.month, dateTo.day)
costResponce = get_costs_by_dates(df, dt)
costs = costResponce[0]['data']

# 
dataxl = createExcel("Розрахунок ", rootPath, df + ' - ' + dt, nameMonth)
prevSheet = dataxl[0]
wb = dataxl[1]
i = 2

users.sort(key=get_department)
for department, data in groupby(users, key=get_department):
    users = list(data)
    if department in departments:
        i += 1
        prevSheet.Cells(i, "B").Value = departments[department]
        prevSheet.Cells(i, "B").Font.Bold = True
        i += 2
        for user in users:
            curUserCosts = list(filter(lambda cost: cost['user_from']['name'] ==
                                user['name'], costs))
            curUserCosts.sort(key=get_projnames)
            if len(curUserCosts) != 0:
                paidCosts = 0
                prevSheet.Cells(i, "A").Value = user['name']
                realcostsByTask = 0
                for paidCost in curUserCosts:
                    paidCosts += float(paidCost['money'])
                sal = round(paidCosts * 100 / (workingDays*8), 1)
                prevSheet.Cells(i, "D").Value = "{}%".format(sal)
                prevSheet.Cells(i, "E").Value = 0
                prevSheet.Cells(i, "F").Formula = "=D{}*E{}".format(i, i)
                i += 1

wb.Save()
wb.Close(True)
print('done')