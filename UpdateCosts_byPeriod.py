from itertools import groupby
from funct import get_users, get_costs_by_dates, get_costs, toDecimal, update_costs, get_tasks, departments


def searchNControl(tasksByProject):                    
    for task in tasksByProject:
        if 'child' in task.keys():
            searchNControl(task['child'])
        else:
            if task['status'] == 'active' and 'tags' in task.keys() and "676734" in task['tags']:
                if task['user_to']['name'] == user['name']:
                    global errorFromNControl
                    errorFromNControl = True
                    break


def get_projnames(iterable):
    return [iterable['task']['project']['name'],
            iterable['task']['project']['page']]


def get_taskname(iterable): 
    return [iterable['task']['name'],
            iterable['task']['id'],
            iterable['task']['page']]
             

# отримання списка співробітників та їх витрат
users = get_users()['data']
costResponce = get_costs_by_dates("'2024-01-01'", "'2024-01-31'")
byPerson = True
personName = 'Горбуля Віталій'
costs = costResponce[0]['data']
costsByProjects = []
tasksByProjects = {}

if byPerson:
    users = list(filter(lambda user: user['name'] ==
                        personName, users))
    curUserCosts = list(filter(lambda cost: cost['user_from']['name'] ==
                            personName, costs))
    for projname, data in groupby(curUserCosts, key=get_projnames):
        if projname[0] not in tasksByProjects.keys():
            tasksByProjects[projname[0]] =  get_tasks(projname[1])['data']
else:
    for projname, data in groupby(costs, key=get_projnames):
        if projname[0] not in tasksByProjects.keys():
            tasksByProjects[projname[0]] =  get_tasks(projname[1])['data']
for user in users:
    if user['department'] in departments:
        curUserCosts = list(filter(lambda cost: cost['user_from']['name'] ==
                            user['name'], costs))
        curUserCosts.sort(key=get_projnames)
        groupedCurUserCostsByProject = {}
        for projname, data in groupby(curUserCosts, lambda cost: cost['task']['project']['name']):
            groupedCurUserCostsByProject[projname] = list(data)
        for projname in groupedCurUserCostsByProject:
            costsByProject = groupedCurUserCostsByProject[projname]

            # перевірка на нормоконтроль
            tasksByProject = tasksByProjects[projname]
            errorFromNControl = False
            searchNControl(tasksByProject)

            if not errorFromNControl:
                costsByProject.sort(key=get_taskname)
                for taskname, costsByTask in groupby(costsByProject, key=get_taskname):
                    listCostsByTask = list(costsByTask)
                    planCost, allCostsbyTask, realCost, = 0, 0, 0
                    # оновлюємо кости
                    taskNControl = False
                    try:
                        if "676734" in listCostsByTask[0]['task']['tags']:
                            for cost in listCostsByTask:
                                update_costs(cost['task']['page'], cost['id'], 0)
                            taskNControl = True
                    except:
                        pass
                    if not taskNControl:
                        taskOnThirdLevel = False
                        try:
                            if listCostsByTask[0]['task']['parent']['parent']:
                                allCostsbyTask = toDecimal(get_costs(taskname[2])['total']['time'])
                                taskOnThirdLevel = True
                        except KeyError:
                                for cost in listCostsByTask:
                                    update_costs(cost['task']['page'], cost['id'], 0)
                        if taskOnThirdLevel:
                            for cost in listCostsByTask:
                                # 675509 : 1:1   - це тег яким помічені всі задачі які мають бути оплачені по факту
                                try:
                                    if cost['task']['tags']:
                                        if "675509" in cost['task']['tags']:
                                            for cost in listCostsByTask:
                                                update_costs(cost['task']['page'], cost['id'], toDecimal(cost['time']))                                        
                                            break
                                except:
                                    pass
                                if cost['task']['status'] == 'done':
                                    try:
                                        planCost = toDecimal(cost['task']['max_time'])
                                    except KeyError:
                                        planCost = 0
                                    try:
                                        realCost = float(cost['money'])
                                        if allCostsbyTask == 0:
                                            k = 0
                                        else:
                                            k = planCost/allCostsbyTask
                                        realCost = toDecimal(cost['time']) * k
                                        update_costs(cost['task']['page'], cost['id'], realCost)
                                    except KeyError:
                                        None
                                else:
                                    update_costs(cost['task']['page'], cost['id'], 0)
            else:
                for cost in costsByProject:
                    update_costs(cost['task']['page'], cost['id'], 0)
        print (user['name'] + " оновлено")
print ('done updated costs')