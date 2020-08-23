# -*- coding: utf-8 -*-

import os, time, sys,inspect
from random import choice
from datetime import datetime,timedelta,date
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase

reload(sys)
sys.setdefaultencoding('utf-8')


textboxBrowse = "D:\Tests\profiles\Тестовая папка".decode('utf-8')
folderName=textboxBrowse
mydir=textboxBrowse
filename = inspect.getframeinfo(inspect.currentframe()).filename
path_dir_root=os.path.dirname(os.path.abspath(filename)).replace("venv\MyOfficeV5", "")
path_dirname_ = os.path.dirname(os.path.abspath(filename)).replace("venv\MyOfficeV5", u"шаблоны")



mydir0 = mydir + "\\" + u"Ошибки"
mydir1 = mydir + "\\" + u"Результаты"
mydir2 = mydir + "\\" + u"Грамоты"
mydir3 = path_dirname_ + "\\" + u"Пример Диплома_Имен. падеж_27.07VMyoffice.docx"
mydir4 = path_dirname_ + "\\" + u"3.1. информац. сообщение детям.docx"
mydir5 = path_dirname_ + "\\" + u"3.2. информац. сообщение детям.docx"
mydir6 = mydir + "\\" + u"Анкеты"
mydir7 = path_dirname_ + "\\" + u"Сводный файл (с учетом новой анкеты участника)_новый.xlsx"
mydir8= path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_исправленный.xlsx"
mydir9 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_исправленный.xlsx"
mydir10 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_исправленный.xlsx"
mydir11 = mydir + "\\" + os.path.basename(mydir7)

mydir12_out = mydir + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_результаты.xlsx"
mydir13_out = mydir + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_результаты.xlsx"
mydir14_out = mydir + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_результаты.xlsx"
mydirs_ = [mydir0, mydir1, mydir2, mydir3, mydir4, mydir5, mydir6, mydir7,mydir8,mydir9,mydir10, mydir11,mydir12_out,mydir13_out,mydir14_out]

global application
application = sdk.Application()
document_xls = application.loadDocument((path_dirname_+"\\"+u"№_1_Егорова_Дарья.xlsx").encode('utf8'))
table_output_xlsx = document_xls.getBlocks().getTable(0)
first_names_F=["Дарья","Ольга",
             "Анастасия", "Анна", "Мария",
             "Алина","Ирина","Екатерина","Арина"]
first_names_M=[
"Алан",
"Александр",
"Алексей",
"Альберт",
"Анатолий",
"Андрей",
"Антон",
"Арсен",
"Арсений",
"Артем",
"Артемий",
"Артур",
"Богдан",
"Борис",
"Вадим",
"Валентин",
"Валерий",
"Василий",
"Виктор",
"Виталий",
"Владимир",
"Владислав",
"Всеволод",
"Вячеслав",
"Геннадий",
"Георгий",
"Герман",
"Глеб",
"Гордей",
]
last_names_M=["Иванов",
              "Смирнов",
              "Кузнецов",
              "Попов", "Васильев", "Петров", "Соколов", "Михайлов", "Новиков", "Фёдоров", "Морозов", "Волков", "Алексеев", "Лебедев", "Семенов", "Егоров", "Павлов", "Козлов"]
middle_name_F=[
"Александровна",
"Алексеевна",
"Анатольевна",
"Андреевна",
"Антоновна",
"Аркадьевна",
"Артемовна",
"Богдановна",
"Борисовна",
"Валентиновна",
"Валерьевна",
"Васильевна",
"Викторовна",
"Виталиевна",
"Владимировна",
"Владиславовна",
]
middle_name_M=[
"Александрович",
"Алексеевич",
"Анатольевич",
"Андреевич",
"Антонович",
"Аркадьевич",
"Арсеньевич",
"Артемович",
"Афанасьевич",
"Богданович",
"Борисович",
"Вадимович",
"Валентинович",
"Валериевич",
"Васильевич",
"Викторович",
"Витальевич",
"Владимирович",
"Всеволодович",
"Вячеславович",
]
sex=["F","M"]
year="01.01."
last_names_F=[i+'а' for i in last_names_M]


for i in range(1,301):
    sex_list = choice(sex)
    date_ = year + str(choice(range(1999, 2010)))
    if sex_list == "F":
        table_output_xlsx.getCell("C8").setText(choice(first_names_F))
        table_output_xlsx.getCell("C6").setText(choice(last_names_F))
        table_output_xlsx.getCell("C10").setText(choice(middle_name_F))
        table_output_xlsx.getCell("C12").setFormattedValue(date_)
    else:
        table_output_xlsx.getCell("C8").setText(choice(first_names_M))
        table_output_xlsx.getCell("C6").setText(choice(last_names_M))
        table_output_xlsx.getCell("C10").setText(choice(middle_name_M))
        table_output_xlsx.getCell("C12").setFormattedValue(date_)
    table_output_xlsx.getCell("C28").setText(choice(["Да","Нет"]))
    table_output_xlsx.getCell("C29").setText(choice(["Да","Нет"]))
    table_output_xlsx.getCell("C30").setText(choice(["Да","Нет"]))
    table_output_xlsx.getCell("C31").setText(choice(["Да","Нет"]))
    table_output_xlsx.getCell("C32").setText(choice(["Да","Нет"]))
    document_xls.saveAs((mydirs_[6] + "\\" + "Анкета "+str(i)+".xlsx").encode('utf8'))



