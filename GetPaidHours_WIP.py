import win32com.client as win32
from funct import get_tasks, get_costs, get_tasktime2level, get_projects, toDecimal

# створюємо excel файл та зберігаємо його
xlApp = win32.Dispatch('Excel.application')
xlApp.Visible = True
wb = xlApp.Workbooks.Add()
prevSheet = wb.Worksheets(1)
project_name = '391_VasylkivLyceum_DD'
prevSheet.Name = project_name[0:30]
prevSheet.Columns("A").ColumnWidth = 35

# 
projects = get_projects()['data']
project = list(filter(lambda projects: projects['name'] == project_name, projects))[0]

# 
prtasks = get_tasks(project['page'])
trackingTimes = get_costs(project['page'])
data_trackingTimes = trackingTimes['data']

costsGeneralByProject = trackingTimes['total']['time']
sortedPrtasks = prtasks['data']
sortedPrtasks.sort(key=lambda x: x.get('name'))

# заповнення заголовку та форматування excel
TasksCount = len(sortedPrtasks)
prevSheet.Cells(1, "A").Value = "Фінансовий Розділ"
prevSheet.Cells(1, "B").Value = "План, год"
prevSheet.Cells(1, "C").Value = "Факт, год"
prevSheet.Cells(TasksCount+2, "A").Value = "Загалом"
prevSheet.Columns("A").ColumnWidth = 35
prevSheet.Columns("B").ColumnWidth = 10
prevSheet.Columns("C").ColumnWidth = 10
prevSheet.Columns("C").NumberFormatLocal = "0,0"
prevSheet.Range("A1:C1").Font.Bold = True
prevSheet.Rows(str(TasksCount+2)).Font.Bold = True
prevSheet.Rows(str(TasksCount+2)).NumberFormatLocal = "0,0"
prevSheet.Cells(str(TasksCount+2), 'C').Value = toDecimal(costsGeneralByProject)

# заповнення даних задач в листі поточного проекта
i = 2
for task in sortedPrtasks:
    prevSheet.Cells(i, "A").Value = task['name']
    try:
        prevSheet.Cells(i, "B").Value = task['max_time']
    except KeyError:
        prevSheet.Cells(i, "B").Value = '0'
    try:
        prevSheet.Cells(i, "C").Value = get_tasktime2level(task['id'], trackingTimes['data'])
    except KeyError:
        prevSheet.Cells(i, "C").Value = 0
    i += 1

# додаємо формули
prevSheet.Cells(str(TasksCount+2), 'B').FormulaLocal = f"=СУММ(B2:B{TasksCount+1})"

