# -*- coding: utf-8 -*-
import os, time,itertools
from random import choice
from datetime import datetime
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase


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

def load_doc(mydirs_):
    input_file_url_xls = mydirs_.encode('utf-8')  # Output с размещением системы рейтингования
    document_xls = application.loadDocument(input_file_url_xls)  # load Сводный
    table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный
    return table_output_xlsx,document_xls



def write_color_win(worker,mydirs_):
    worker.ReportProgress(91, u"Выделение победителя первого этапа")
    for i in range(1, 5):
        globals()['table_output_xlsx_%d' % i],globals()['document_xls_%d' % i] =load_doc(mydirs_[i+10].encode('utf-8'))
    n_rows = table_output_xlsx_1.getRowsCount()
    for i in range(4, n_rows + 1):
        n_rows = str(i)
        if float(table_output_xlsx_1.getCell("W" + n_rows).getFormattedValue()) > 10:
            cell_range = table_output_xlsx_1.getCellRange("B" + n_rows + ":AJ" + n_rows)
            for i in range(1, 5):
                globals()['cell_range_%s' % i] = globals()['table_output_xlsx_%s' % i].getCellRange(
                    "A" + n_rows + ":F" + n_rows)
                globals()['cell_range_%s' % i].setCellProperties(cell_properties)
            cell_range.setCellProperties(cell_properties)
    for i in range(1, 5):
        try:
            globals()['document_xls_%s' % i].saveAs(mydirs_[i+10].encode('utf-8'))
        except Exception as e:
            raise Exception(u"Открыт документ: " + os.path.basename(mydirs_[i+10]))
    worker.ReportProgress(100, u"Завершено")

def write_scores(col,scores_,table_output_xlsx_main):
    n_rows = table_output_xlsx_main.getRowsCount()
    for i in range(4, n_rows + 1):
        table_output_xlsx_main.getCell(col + str(i)).setNumber(scores_[i - 4])
    #document_xls.saveAs(mydirs_.encode('utf8'))


def get_scores(table_output_xlsx,adr):
    scores_ = []
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4,n_rows+1):
        scores_.append(int(table_output_xlsx.getCell(str(adr)+str(i)).getFormattedValue()))
    return scores_




def main_score(worker,mydirs_,adr,k,adr_last,proc):
    table_output_xlsx_main, document_xls_main = load_doc(mydirs_[11])
    worker.ReportProgress(0, u"Экспорт балов.")
    column = list_xls(str(adr_last))
    j = 0  # Смещение вправо
    for i in range(12, 15):
        table_output_xlsx,document_xls=load_doc(mydirs_[i])
        scores_ = get_scores(table_output_xlsx,adr)                  # Экспорт баллов
        write_scores(column[k + j], scores_, table_output_xlsx_main) #Сохранение в сводный по трем файлам, в цикле
        j += 1
        worker.ReportProgress(30*j, u"Экспорт балов.")
    try:
        document_xls_main.saveAs(mydirs_[11].encode('utf8'))
    except Exception as e:
        raise Exception(u"Открыт документ: " + os.path.basename(mydirs_[11]))
    worker.ReportProgress(proc, u"Экспорт завершен.")


