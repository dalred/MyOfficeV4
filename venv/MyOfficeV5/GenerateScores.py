# -*- coding: utf-8 -*-
import os, time,itertools,inspect,re
from random import choice
from datetime import datetime,timedelta,date
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase

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
cell_properties = sdk.CellProperties()
cell_properties.backgroundColor = sdk.ColorRGBA(193, 242, 17, 255)

def iter_all_strings():
    for size in itertools.count(1):
        for s in itertools.product(ascii_uppercase, repeat=size):
            yield "".join(s)

def list_xls(rang):
    lst_addr = []
    for s in iter_all_strings():
        lst_addr.append(s)
        if s == rang:
            break
    return lst_addr

scores_=range(1,5)
scores__=range(1,10)
#datetime=range()  #mm/dd/year

#def Generate_FIO_Date():


def generate_scores(mydir):
    document_xls = application.loadDocument(mydir.encode('utf8'))
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4, n_rows + 1):
        table_output_xlsx.getCell("C"+ str(i)).setNumber(choice(scores_))
        table_output_xlsx.getCell("D" + str(i)).setNumber(choice(scores_))
        table_output_xlsx.getCell("E" + str(i)).setNumber(choice(scores_))
    document_xls.saveAs(mydir.encode('utf8'))

def generate_scores2(mydir):
    document_xls = application.loadDocument(mydir.encode('utf8'))
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4, n_rows + 1):
       if table_output_xlsx.getCell("A" + str(i)).getCellProperties().backgroundColor.__eq__(cell_properties.backgroundColor):
                table_output_xlsx.getCell("A" + str(i)).setNumber(choice(scores__))
                table_output_xlsx.getCell("G" + str(i)).setNumber(choice(scores__))
                table_output_xlsx.getCell("H" + str(i)).setNumber(choice(scores__))
                table_output_xlsx.getCell("I" + str(i)).setNumber(choice(scores__))
                table_output_xlsx.getCell("J" + str(i)).setNumber(choice(scores__))
    document_xls.saveAs(mydir.encode('utf8'))




column = list_xls("AL")
k=29 #AD AH 33
j=0 #Смещение вправо
for i in range(12, 15):
    #generate_scores(mydirs_[i])
    #generate_scores2(mydirs_[i])
    #scores_1,scores_2=get_scores(mydirs_[i])
    #write_scores(column[k+j],column[k+4+j], scores_1, scores_2)
    j+=1

regex_count = '^\d+$'
count=200
if not re.search(regex_count, str(count)):
    print "hello"

