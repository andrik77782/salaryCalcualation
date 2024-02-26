from itertools import groupby
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import os
from funct import get_users, departments

# основні параметри
temp_folder = "C:\\WS, bills\\temp"
subject = 'Нарахування за січень'


# Параметри електронної пошти
smtp_server = 'smtp.gmail.com'
port = 587 
sender_email = 'slyusarenko@aimm-group.com'
password = '$Aimm2020$'

#витягуємо файли з розрахунками
file_list = []
for root_path, directories, files in os.walk(temp_folder):
    for file in files:
        full_path = os.path.join(root_path, file)
        file_list.append(full_path)

#        
contacts = get_users()['data']
contacts.sort(key=lambda x: x['department'])        
groupedContacts = {}
for department, data in groupby(contacts, key=lambda x: x['department']):
    if department in departments:
        groupedContacts[department] = list(data)
#

subordinates = {
    'Прокоп\'юк Олександр' : ['03.1_AR_Група АР1','02_SP_Група ГП'],
    'Загарюк Роман' : ['03.2_AR_Група АР2'],
    'Магера Олександр' : ['03.3_AR_Група АР3'],
    'Іоненко Дмитро' : ['06.1_ST_Група КР1'],
    'Дзюба Тетяна' : ['09.1_HV_Група ОВ1'],
    'Горбуля Віталій' : ['09.2_HV_Група ОВ2'],
    'Токарчук Константин' : ['08.1_WS_Група ВК1'],
    'Білокрис Оксана' : ['11_ED_Група ЕТР1', '35_LC_Група СМ1'],
    'Хмель Людмила' : ['08.2_WS_Група ВК2']
}
for contact in contacts:
    for file in file_list:
        if contact['name'] in os.path.basename(file):
            contact['file_list'] = [file]

for headOfDepartment in contacts:
    if headOfDepartment['name'] in subordinates:
        file_list = []
        for department in subordinates[headOfDepartment['name']]:
            for contact in groupedContacts[department]:
                try:
                    file_list.append(contact['file_list'][0])
                except KeyError:
                    contact['file_list'] = [temp_folder+'\\'+departments[department]+'\\'+contact['name']+' '+subject]
                    file_list.append(contact['file_list'][0])
        headOfDepartment['file_list'] = file_list


for department in groupedContacts:
    for contact in groupedContacts[department]:
        receiver_email = contact['email']
        # Створення повідомлення
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = receiver_email
        message['Subject'] = subject
        body = ''
        message.attach(MIMEText(body, 'plain'))
        if 'file_list' in contact:
        # додаємо вкладення
            for file_mail in contact['file_list']:
                if os.path.exists(file_mail) == False:
                    continue
                attachment = open(file_mail, 'rb')
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_mail))
                message.attach(part)
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()  # Зашифровуємо з'єднання TLS
            server.ehlo()
            server.login(sender_email, password)
            text = message.as_string()
            server.sendmail(sender_email, receiver_email, text)
            print(receiver_email)

print('done')