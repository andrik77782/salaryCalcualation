import hashlib
import requests
import psycopg2

ApiKey = "fbea30a735ed774f6814354712f20b65"

departments = {'01_CD_Креативна група': 'Креатив',
                '01_ID_Група АІ': 'АІ',
                '02_SP_Група ГП': 'ГП',
                '03.1_AR_Група АР1': 'АР1',
                '03.2_AR_Група АР2': 'АР2',
                '03.3_AR_Група АР3': 'АР3',
                '06.1_ST_Група КР1': 'КР1',
                '06.2_ST_Група КР2': 'КР2',
                '08.1_WS_Група ВК1': 'ВК1',
                '08.2_WS_Група ВК2': 'ВК2',
                '11_ED_Група ЕТР1': 'ЕТР1',
                '09.1_HV_Група ОВ1': 'ОВ1',
                '09.2_HV_Група ОВ2': 'ОВ2',
                '00_PMO_Офіс управління проектами' : 'PMO',
                '35_LC_Група СМ1': 'СМ1'}


def Hash(toHash):
    he = toHash.encode('utf-8')
    hh = hashlib.md5(he)
    return hh.hexdigest()


def get_events(period):
    query = ("https://aimm.worksection.com/api/admin/v2/?action=get_events&period=" + 
             period + '&hash=' + Hash("get_events"+ApiKey))
    return requests.get(query).json() 


def get_users():
    query = ("https://aimm.worksection.com/api/admin/v2/?action=get_users&hash=" + Hash("get_users"+ApiKey))
    return requests.get(query).json()


def update_costs(page, costId, money):
    query = ("https://aimm.worksection.com/api/admin/v2/?action=update_costs&page=" +
             page + "&id=" + costId + "&money=" + str(money) + "&hash=" + Hash(page +
             "update_costs"+ApiKey))
    req = requests.get(query).json()
    return None


def update_task_dates(page, datestart, dateend):
    query = ("https://aimm.worksection.com/api/admin/v2/?action=update_task&page=" +
             page + "&datestart=" + datestart + "&dateend=" + dateend + "&hash=" + Hash(page +
             "update_task"+ApiKey))
    req = requests.get(query).json()
    return None
    

def get_tags():
    query = ("https://aimm.worksection.com/api/admin/v2/?action=get_tags&" +
             "&hash=" + Hash("get_tags"+ApiKey))
    return requests.get(query).json()


def get_tag_groups():
    query = ("https://aimm.worksection.com/api/admin/v2/?action=get_tag_groups&" +
             "&hash=" + Hash("get_tag_groups"+ApiKey))
    return requests.get(query).json()


def get_tasks(Page):
    query = ("https://aimm.worksection.com/api/admin/v2/?action=get_tasks&page=" +
             Page + "&hash=" + Hash(Page+"get_tasks"+ApiKey) + "&show_subtasks=1")
    return requests.get(query).json()


def get_activeTasks(Page):
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_tasks&page=" + Page + "&hash=" + Hash(Page+"get_tasks"+ApiKey) + "&show_subtasks=1&filter=active"
    return requests.get(query).json()


def get_costs(Page):
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_costs&page=" + Page + "&hash=" + Hash(Page+"get_costs"+ApiKey) 
    return requests.get(query).json()


def get_costs_by_dates(dateFrom, dateTo):
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_costs&hash=" + Hash("get_costs"+ApiKey) + "&filter=dateadd>=" + dateFrom + " and dateadd<=" + dateTo

    return [requests.get(query).json(), dateFrom, dateTo]


def get_projects():
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_projects&hash=" + Hash("get_projects"+ApiKey) + "&filter=active"
    return requests.get(query).json()


def get_projects(extra):
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_projects&hash=" + Hash("get_projects"+ApiKey) + "&filter=active&extra=" + extra
    return requests.get(query).json()


def get_users():
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_users&hash=" + Hash("get_users"+ApiKey)
    return requests.get(query).json()


def get_comments(Page):
    query = "https://aimm.worksection.com/api/admin/v2/?action=get_comments&page=" + Page + "&hash=" + Hash(Page+"get_comments"+ApiKey)
    return requests.get(query).json()


def toDecimal(time):
    try:
        (h, m) = time.split(':')
    except AttributeError:
        return time
    return (int(h) * 60 + int(m))/60


def get_tasktime2level(id, tasks):
    time = 0
    for item in tasks:
        if id == item['task']['id']:
            if item['time']:
                time += toDecimal(item['time'])
    for item in tasks:
        try:
            if id == item['task']['parent']['id']:
                if item['time']:
                    time += toDecimal(item['time'])
        except KeyError:
            time += 0
    return time


def get_tasktime(id, tasks):
    time = 0
    for item in tasks:
        if id == item['task']['id']:
            if item['time']:
                time += toDecimal(item['time'])
    return time

# create connection to database
def open_connection_to_salary_db():
    server = '127.0.0.1'
    database = 'salary_db'
    username = 'postgres'
    password = '14789632'
    conn = psycopg2.connect("host={} dbname={} user={} password={}".format(server, database, username, password))
    conn.set_session(autocommit=True)
    cursor = conn.cursor()
    return conn,cursor