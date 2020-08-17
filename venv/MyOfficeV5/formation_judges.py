# -*- coding: utf-8 -*-
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
import inspect, os

application = sdk.Application()



def load_doc(mydirs_):
    input_file_url_xls = mydirs_.encode('utf-8')  # Output с размещением системы рейтингования
    document_xls = application.loadDocument(input_file_url_xls)  # load Сводный
    table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный
    return table_output_xlsx,document_xls


def write_formules(row_str,fio_list,table_output2_xlsx,table_output_xlsx,mydirs_,worker):
    worker.ReportProgress(10, u"Формирование файлов Жюри")
    current_row = 4 # c index 4
    for str_ in fio_list:
        index = current_row - 3
        percentage = (index * 95) / len(fio_list)
        worker.ReportProgress(percentage, u"Формирование файлов Жюри")
        row_str = str(current_row)
        table_output2_xlsx.getCell("B" + row_str).setText(str_)
        table_output2_xlsx.getCell("A" + row_str).setNumber(index)
        table_output2_xlsx.getCell("F" + row_str).setFormula("=SUM(C" + row_str + ":D" + row_str + ":E" + row_str + ")")
        table_output2_xlsx.getCell("K" + row_str).setFormula("=SUM(G" + row_str + ":H" + row_str + ":I" + row_str + ":J" + row_str + ")")
        table_output_xlsx.getCell("W" + row_str).setFormula("=AVERAGE(AC" + row_str + ":AE" + row_str + ")+SUM(" + "X" + row_str + ":AB" + row_str + ")")
        table_output_xlsx.getCell("AF" + row_str).setFormula("=AVERAGE(AG" + row_str + ":AI" + row_str + ")")
        table_output_xlsx.getCell("AJ" + row_str).setFormula("=SUM(AF" + row_str + ",W" + row_str + ")")
        current_row += 1


def main_judges(worker,mydirs_):
    worker.ReportProgress(0, u"Формирование файлов Жюри")
    table_output_xlsx,document_xls =load_doc(mydirs_[11])
    table_output2_xlsx,document2_xls=load_doc(mydirs_[8])
    row_str = table_output_xlsx.getRowsCount()  # количество строк
    E4_empty = table_output_xlsx.getCell("E4").getRawValue() == ''

    if not E4_empty:
        cell_range = table_output_xlsx.getCellRange("E4:E" + str(row_str))
        fio_list = [i.getRawValue() for i in cell_range]

    # Выделение строк в таблице по количеству строк из документов
    rows_count_2 = len(fio_list)
    if rows_count_2 > 1:
        table_output2_xlsx.insertRowAfter(3, copyRowStyle=True, rowsCount=rows_count_2 - 1)
        document2_xls.saveAs(mydirs_[12].encode('utf-8'))
    elif rows_count_2 == 1:
        document2_xls.saveAs(mydirs_[12].encode('utf-8'))

    write_formules(row_str,fio_list,table_output2_xlsx,table_output_xlsx,mydirs_,worker)
    document_xls.saveAs(mydirs_[11].encode('utf-8'))
    document2_xls.saveAs(mydirs_[12].encode('utf-8'))
    table_output2_xlsx.getCell("B1").setText("Жюри 2")
    document2_xls.saveAs(mydirs_[13].encode('utf-8'))
    table_output2_xlsx.getCell("B1").setText("Жюри 3")
    document2_xls.saveAs(mydirs_[14].encode('utf-8'))
    worker.ReportProgress(100, u"Завершено формирование файлов Жюри")



