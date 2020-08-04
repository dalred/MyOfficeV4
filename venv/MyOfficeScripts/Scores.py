# -*- coding: utf-8 -*-
import os, time,itertools
from random import choice
from datetime import datetime
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase


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


def write_scores(col1,col2,scores1,scores2,mydirs_):
    document_xls=application.loadDocument(mydirs_.encode('utf8'))
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4, n_rows + 1):
        table_output_xlsx.getCell(col1 + str(i)).setNumber(scores1[i - 4])
        table_output_xlsx.getCell(col2 + str(i)).setNumber(scores2[i - 4])
    document_xls.saveAs(mydirs_.encode('utf8'))


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


def main_score(worker,mydirs_):
    worker.ReportProgress(0, u"Экспорт балов.")
    column = list_xls("AL")
    k = 29  # AD AH 33
    j = 0  # Смещение вправо
    for i in range(12, 15):
        # generate_scores(mydirs_[i])
        scores_1, scores_2 = get_scores(mydirs_[i])
        write_scores(column[k + j], column[k + 4 + j], scores_1, scores_2,mydirs_[11])
        j += 1
        worker.ReportProgress(30*j, u"Экспорт балов.")
    worker.ReportProgress(100, u"Экспорт завершен.")

