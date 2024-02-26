import openpyxl
from funct import get_projects, get_tasks, update_task_dates

def read_excel_file(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    worksheet = workbook.active
    data = []
    for row in worksheet.iter_rows(min_row=9, values_only=True):
        data.append({"task_name": row[4], "start_date": row[5], "end_date": row[8]})
    return data


file_path = "C:\\Users\\andri\\Downloads\\Telegram Desktop\\2024-02-19_Project Schedule N01_FoS_Fosa.xlsx"
excel_data = read_excel_file(file_path)
tasks_data = get_tasks("/project/284309/")['data']
flatten_tasks_data = [{task['name']: task['page']} for task in tasks_data]
for task in tasks_data:
    if 'child' in task.keys():
        for child in task['child']:
            flatten_tasks_data.append({child['name']: child['page']})
nums = list()
for task in flatten_tasks_data:
    nums.append(len(task))

for task_data in excel_data:
    task_name = task_data["task_name"]
    if task_name:
        task_name = task_name.strip()
    start_date = task_data["start_date"]
    end_date = task_data["end_date"]

    if task_name is not None and start_date is not None and end_date is not None:
        # Знаходимо задачі, які відповідають назві задачі з Excel
        matching_task = list(filter(lambda elem: task_name in elem.keys(), flatten_tasks_data))
        
        if matching_task:
            # Оновлюємо дати для знайдених задач у Worksection
            update_task_dates(matching_task[0][task_name], start_date.strftime("%d.%m.%Y"), end_date.strftime("%d.%m.%Y"))
            print(f"Оновлено дати для задачі '{task_name}")
        else:
            print(f"Не знайдено задачі '{task_name}")
    else:
        print(f"Відсутні дані для задачі '{task_name}' від {start_date} до {end_date}")

