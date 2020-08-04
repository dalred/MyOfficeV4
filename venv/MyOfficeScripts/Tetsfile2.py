# -*- coding: utf-8 -*-
from re import compile
import os, time,itertools,inspect
from random import choice
from datetime import datetime
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase

textboxBrowse = "D:\Tests\profiles\Тестовая папка".decode('utf-8')
folderName=textboxBrowse
filename = inspect.getframeinfo(inspect.currentframe()).filename
path_dirname_ = os.path.dirname(os.path.abspath(filename)).replace("\\venv\\MyOfficeV3", u"\\шаблоны")




mydir0 = textboxBrowse + "\\" + u"Ошибки"
mydir1 = textboxBrowse + "\\" + u"Результаты"
mydir2 = textboxBrowse + "\\" + u"Грамоты"
mydir3 = path_dirname_ + "\\" + u"Сводный файл (с учетом новой анкеты участника)_30.07.2020.xlsx"
mydir4 = path_dirname_ + "\\" + u"Пример Диплома_Имен. падеж_27.07VMyoffice.docx"
mydir5 = path_dirname_ + "\\" + u"3.1. информац. сообщение детям.docx"
mydir6 = path_dirname_ + "\\" + u"3.2. информац. сообщение детям.docx"
mydir7 = textboxBrowse + "\\" + u"Анкеты"
mydir8= path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_исправленный.xlsx"
mydir9 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_исправленный.xlsx"
mydir10 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_исправленный.xlsx"

mydir11 = textboxBrowse + "\\" + os.path.basename(mydir3)
mydir12_out = textboxBrowse + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_результаты.xlsx"
mydir13_out = textboxBrowse + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_результаты.xlsx"
mydir14_out = textboxBrowse + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_результаты.xlsx"
mydirs_ = [mydir0, mydir1, mydir2, mydir3, mydir4, mydir5, mydir6, mydir7,mydir8,mydir9,mydir10, mydir11,mydir12_out,mydir13_out,mydir14_out]





global application
application = sdk.Application()

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

def generate_scores(mydir):
    document_xls = application.loadDocument(mydir.encode('utf8'))
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4, n_rows + 1):
        table_output_xlsx.getCell("C"+ str(i)).setNumber(choice(scores_))
        table_output_xlsx.getCell("D" + str(i)).setNumber(choice(scores_))
        table_output_xlsx.getCell("E" + str(i)).setNumber(choice(scores_))
        table_output_xlsx.getCell("G" + str(i)).setNumber(choice(scores__))
        table_output_xlsx.getCell("H" + str(i)).setNumber(choice(scores__))
        table_output_xlsx.getCell("I" + str(i)).setNumber(choice(scores__))
        table_output_xlsx.getCell("J" + str(i)).setNumber(choice(scores__))
    document_xls.saveAs(mydir.encode('utf8'))

def write_scores(col1,col2,scores1,scores2):
    document_xls=application.loadDocument(mydirs_[11].encode('utf8'))
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4, n_rows + 1):
        table_output_xlsx.getCell(col1 + str(i)).setNumber(scores1[i - 4])
        table_output_xlsx.getCell(col2 + str(i)).setNumber(scores2[i - 4])
    document_xls.saveAs(mydirs_[11].encode('utf8'))



def get_scores(mydir):
    scores_1 = []
    scores_2 = []
    document_xls = application.loadDocument(mydir.encode('utf8'))
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4,n_rows+1):
        scores_1.append(int(table_output_xlsx.getCell('F'+str(i)).getFormattedValue()))
        scores_2.append(int(table_output_xlsx.getCell('K'+str(i)).getFormattedValue()))
    return scores_1,scores_2





column = list_xls("AL")
k=29 #AD AH 33
j=0 #Смещение вправо
for i in range(12, 15):
    generate_scores(mydirs_[i])
    """scores_1,scores_2=get_scores(mydirs_[i])
    write_scores(column[k+j],column[k+4+j], scores_1, scores_2)
    j+=1"""

