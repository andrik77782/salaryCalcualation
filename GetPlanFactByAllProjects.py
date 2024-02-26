import win32com.client as win32
from funct import get_tasks, get_costs, get_tasktime2level, get_projects, toDecimal

# створюємо excel файл та зберігаємо його
xlApp = win32.Dispatch('Excel.application')
xlApp.Visible = True
wb = xlApp.Workbooks.Add()
    # wb.SaveAs(os.path.join(os.getcwd(), 'test.xlsx'))
prevSheet = wb.Worksheets(1)
prevSheet.Name = 'Список проектів'
prevSheet.Columns("A").ColumnWidth = 35

# отримуємо і зберігаємо перелік активних проектів
projects = get_projects()['data']
projects.sort(key=lambda x: x.get('name'), reverse=True)

# проходимо по списку всіх проектів
j = 1
for project in projects:
    # створюємо лист з назвою проекта, додаємо лінки
    projectSheet = wb.Sheets.Add(After=prevSheet)
    projectSheet.Name = project['name'][0:30]
    backShape = projectSheet.Shapes.AddShape(133, 300, 10, 40, 20)
    projectSheet.Hyperlinks.Add(Anchor=backShape,
                                SubAddress="'{}'!A1".format(prevSheet.Name),
                                Address="",
                                TextToDisplay="Список проектів",
                                ScreenTip="Список проектів")
    prevSheet.Hyperlinks.Add(Anchor=prevSheet.Range('A{}'.format(j)),
                            SubAddress="'{}'!A1".format(projectSheet.Name),
                            Address="",
                            TextToDisplay=projectSheet.Name,
                            ScreenTip="Перейти до проекту")

    # отримуємо дані по задачам з worksection та сортуємо по імені розділу
    prtasks = get_tasks(project['page'])
    trackingTimes = get_costs(project['page'])
    costsGeneralByProject = trackingTimes['total']['time']
    sortedPrtasks = prtasks['data']
    sortedPrtasks.sort(key=lambda x: x.get('name'))

    # заповнення заголовку та форматування excel
    TasksCount = len(sortedPrtasks)
    projectSheet.Cells(1, "A").Value = "Фінансовий Розділ"
    projectSheet.Cells(1, "B").Value = "План, год"
    projectSheet.Cells(1, "C").Value = "Факт, год"
    projectSheet.Cells(TasksCount+2, "A").Value = "Загалом"
    projectSheet.Columns("A").ColumnWidth = 35
    projectSheet.Columns("B").ColumnWidth = 10
    projectSheet.Columns("C").ColumnWidth = 10
    projectSheet.Columns("C").NumberFormatLocal = "0,0"
    projectSheet.Range("A1:C1").Font.Bold = True
    projectSheet.Rows(str(TasksCount+2)).Font.Bold = True
    projectSheet.Rows(str(TasksCount+2)).NumberFormatLocal = "0,0"
    projectSheet.Cells(str(TasksCount+2), 'C').Value = toDecimal(costsGeneralByProject)

    # заповнення даних задач в листі поточного проекта
    i = 2
    for task in sortedPrtasks:
        projectSheet.Cells(i, "A").Value = task['name']
        try:
            projectSheet.Cells(i, "B").Value = task['max_time']
        except KeyError:
            projectSheet.Cells(i, "B").Value = '0'
        try:
            projectSheet.Cells(i, "C").Value = get_tasktime2level(task['id'], trackingTimes['data'])
        except KeyError:
            projectSheet.Cells(i, "C").Value = 0
        if projectSheet.Cells(i, "B").Value < projectSheet.Cells(i, "C").Value:
            projectSheet.Cells(i, "C").Interior.ColorIndex = 3
        i += 1

    # додаємо формули
    projectSheet.Cells(str(TasksCount+2), 'B').FormulaLocal = f"=СУММ(B2:B{TasksCount+1})"

    j += 1

