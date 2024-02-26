import win32com.client as win32
import os, datetime
from itertools import groupby
from funct import get_users, get_costs_by_dates, toDecimal, get_tasks, departments


def searchNControl(tasksByProject):                    
    for task in tasksByProject:
        if 'child' in task.keys():
            searchNControl(task['child'])
        else:
            if task['status'] == 'active' and 'tags' in task.keys() and "676734" in task['tags']:
                if task['user_to']['name'] == user['name']:
                    global needNControl
                    needNControl = True
                    break


def get_tasknameid(iterable): 
    return [iterable['task']['name'],
            iterable['task']['id']]


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


def createExcel(userName, path, date, monthName):

    # створюємо excel файл та зберігаємо його
    wb = xlApp.Workbooks.Add()    
    sheet = wb.Worksheets(1)
    sheet.Name = userName
    sheet.Columns("A:B").ColumnWidth = 2.4
    sheet.Columns("C").ColumnWidth = 80
    sheet.Range("A1 : O4").Font.Bold = True
    sheet.Cells(1, "A").Font.Size = 15

    # заповнюємо шапку в ексель
    sheet.Cells(2, "A").Value = date
    sheet.Cells(4, "D").Value = "Факт"
    sheet.Cells(4, "E").Value = "Нараховано"

    return sheet, wb

# xlApp = win32.gencache.EnsureDispatch('Excel.Application')
xlApp = win32.Dispatch('Excel.application')
xlApp.Visible = False

# основні дані
rootPath = "C:\\WS, bills\\temp"
nameMonth = " січень 2024"
dateFrom = datetime.datetime(2024, 1, 1)
dateTo = datetime.datetime(2024, 1, 31)
workingDays = 22
days = (dateTo - dateFrom).days + 1
byPerson = True
personName = 'Горбуля Віталій'


# отримання списка співробітників та їх витрат за певний період
users = get_users()['data']
if byPerson:
    users = list(filter(lambda user: user['name'] ==
                        personName, users))
df = "'{}-{:02}-{:02}'".format(dateFrom.year, dateFrom.month, dateFrom.day)
dt = "'{}-{:02}-{:02}'".format(dateTo.year, dateTo.month, dateTo.day)
costResponce = get_costs_by_dates(df, dt)
costs = costResponce[0]['data']

# отримання списка проектів де були витрати в вибраний період і збереження всіх тасків по ним
tasksByProjects = {}
if byPerson:
    costs = list(filter(lambda cost: cost['user_from']['name'] ==
                         personName, costs))
costs.sort(key=get_projnames)
for projname, data in groupby(costs, key=get_projnames):
    responce = get_tasks(projname[1])
    tasksByProjects[projname[0]] =  responce['data']

costsByProjects = []

for user in users:
    department = user['department']
    if department in departments:
        path = os.path.join(rootPath, departments[department])
        try:
            os.mkdir(path)
        except OSError:
            None
        i = 4
        curUserCosts = [cost for cost in costs if cost['user_from']['name'] == user['name']]
        curUserCosts.sort(key=get_projnames)
        if len(curUserCosts) != 0:
            xl = createExcel(user['name'], path, df + ' - ' + dt, nameMonth)
            prevSheet = xl[0]
            wb = xl[1]
            trackedCosts, paidCosts = 0, 0
            curProjectForPrint = ""
            prevSheet.Cells(1, "A").Value = user['name']
            i += 1

            groupedUserCostsByProject = {}
            for projname, data in groupby(curUserCosts, lambda cost: cost['task']['project']['name']):
                groupedUserCostsByProject[projname] = list(data)

            for projname in groupedUserCostsByProject:
                # беремо список витрат по поточному проекту та сортуємо по задачам
                userCostsByProject = groupedUserCostsByProject[projname]
                userCostsByProject.sort(key=get_taskid)
                # прописуємо назву проекта
                prevSheet.Cells(i, "A").Value = "  " + projname
                prevSheet.Cells(i, "A").Font.Bold = True
                prevSheet.Cells(i, "A").Font.Size = 16
                i += 1
                rowsForWrite = []

                # search for N control
                tasksByProject = tasksByProjects[projname]
                needNControl = False                
                searchNControl(tasksByProject)

                # proccess all costs by task
                for id, costsByTask in groupby(userCostsByProject, key=get_taskid):
                    listCostsByTask = list(costsByTask)
                    costsByTask, realcostsByTask = 0, 0
                    for cost in listCostsByTask:
                        costsByTask += toDecimal(cost['time'])
                        paidCosts += float(cost['money'])
                        realcostsByTask += float(cost['money'])
                        trackedCosts += toDecimal(cost['time'])
                    names = ""
                    br = 0
                    row = []
                    for task in tasksByProject:
                        if br == 1:
                            break
                        else:
                            if task['id'] == id:
                                row.append([task['name']])
                                row.append([task['page']])
                                row.append(task['status'])
                                row.append('level')
                                br = 1
                                break
                            else:
                                if 'child' in task.keys():
                                    names = task['name']
                                    for subtask in task['child']:
                                        if subtask['id'] == id:
                                            row.append([task['name'], subtask['name']])
                                            row.append([task['page'], subtask['page']])
                                            row.append(subtask['status'])
                                            row.append('level')
                                            br = 1
                                            break
                                        else:
                                            if 'child' in subtask.keys():
                                                for sub_task in subtask['child']:
                                                    if sub_task['id'] == id:
                                                        row.append([task['name'], subtask['name'], sub_task['name']])
                                                        row.append([task['page'], subtask['page'], sub_task['page']])
                                                        row.append(sub_task['status'])
                                                        if not needNControl:
                                                            try:
                                                                row.append(sub_task['max_time'])
                                                            except KeyError:
                                                                row.append("no_plan")
                                                        else:
                                                            row.append("ncontrol")
                                                        br = 1
                                                        break
                    row.append(round(costsByTask, 2))
                    row.append(round(realcostsByTask, 2))
                    rowsForWrite.append(row)
                rowsForWrite.sort(key=get_taskname)
                for nameRozdil, row in groupby(rowsForWrite, key=get_taskname):
                    listrow = list(row)                        
                    prevSheet.Cells(i, "A").Value = nameRozdil
                    prevSheet.Hyperlinks.Add(Anchor=prevSheet.Cells(i, "A"),
                                            Address="https://aimm.worksection.com{}".format(listrow[0][1][0]))
                    prevSheet.Cells(i, "A").Font.Size = 14
                    for row in listrow:
                        if len(row[0]) == 1:
                            prevSheet.Cells(i, "D").Value = row[4]
                            prevSheet.Cells(i, "E").Value = row[5]
                            if row[5] in (0, '0.0'):
                                if row[3] == 'level':
                                    prevSheet.Cells(i, "G").Value = "* задача не на 3-му рівні, тому не зараховано"
                                elif row[3] == 'ncontrol':
                                    prevSheet.Cells(i, "G").Value = "* є зауваження експертизи, тому не зараховано"
                            listrow.remove(row)
                            break
                    i += 1
                    k = 0
                    listrow.sort(key=get_taskname2)
                    for nameBlock, row2 in groupby(listrow, key=get_taskname2):
                        listrow2 = list(row2)  
                        prevSheet.Cells(i, "B").Value = nameBlock
                        prevSheet.Hyperlinks.Add(Anchor=prevSheet.Cells(i, "B"),
                                                Address="https://aimm.worksection.com{}".format(listrow2[0][1][1]))
                        prevSheet.Cells(i, "B").Font.Size = 12
                        for row in listrow2:
                            if len(row[0]) == 2:
                                prevSheet.Cells(i, "D").Value = row[4]
                                prevSheet.Cells(i, "E").Value = row[5]
                                if row[3] == 'level':
                                    prevSheet.Cells(i, "G").Value = "* задача не на 3-му рівні, тому не зараховано"
                                elif row[3] == 'ncontrol':
                                    prevSheet.Cells(i, "G").Value = "* є зауваження експертизи, тому не зараховано"
                                listrow2.remove(row)
                                break
                        listrow2.sort(key=get_taskname3)
                        j = 0
                        for row in listrow2:
                            i += 1
                            prevSheet.Cells(i, "C").Value = row[0][2]
                            prevSheet.Hyperlinks.Add(Anchor=prevSheet.Cells(i, "C"),
                                                Address="https://aimm.worksection.com{}".format(listrow2[j][1][2]))
                            prevSheet.Cells(i, "D").Value = row[4]
                            prevSheet.Cells(i, "E").Value = row[5]
                            if row[5] in (0, '0.0'):
                                if row[3] == 'level':
                                    prevSheet.Cells(i, "G").Value = "* задача не на 3-му рівні, тому не зараховано"
                                elif row[3] == 'ncontrol' and row[2] == 'active':
                                    prevSheet.Cells(i, "G").Value = "*  заблоковано, бо є зауваження від нормоконтролю"
                                elif row[2] == 'active':
                                    prevSheet.Cells(i, "G").Value = "* задача не закрита, тому не зараховано"
                                elif row[3] == 'no_plan':
                                    prevSheet.Cells(i, "G").Value = "* не вказан план, тому не зараховано"
                            else:
                                prevSheet.Cells(i, "G").Value = "".format()
                            j += 1
                        i += 1
                        k += 1


            
            # пишемо підсумок по фактичним та перерахованим годинам
            i += 1
            prevSheet.Cells(i, "D").Value = trackedCosts
            prevSheet.Cells(i, "E").Value = paidCosts
            prevSheet.Range("B{} : E{}".format(i, i)).Font.Bold = True
            prevSheet.Range("B{} : E{}".format(i, i)).Font.Size = 14
            i += 2
            sal = round(paidCosts * 100 / (workingDays*8), 1)
            prevSheet.Cells(i, "E").Value = "{} / {}*8 = {}% від ставки".format(round(paidCosts, 1), workingDays, sal)
            prevSheet.Range("B{} : E{}".format(i, i)).Font.Bold = True
            prevSheet.Range("B{} : E{}".format(i, i)).Font.Size = 20
            print(user['name'] +  ' xlsx додано')
            wb.SaveAs(os.path.join(path, '{}{}.xlsx'.format(user['name'], nameMonth)))
            wb.Close(True)
print ('done')