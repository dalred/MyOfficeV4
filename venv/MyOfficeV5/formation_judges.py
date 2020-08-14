# -*- coding: utf-8 -*-
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
import inspect, os


def write_formules(n_rows,fio_list,table_output2_xlsx,table_output_xlsx,document2_xls,mydirs_,worker):
    worker.ReportProgress(10, u"Формирование файлов Жюри")
    current_row = 4 # c index 4
    for i in range(current_row, n_rows + 1):
        index = i - 3
        n_rows = str(i)
        percentage = (index * 100) / len(fio_list)
        worker.ReportProgress(percentage, u"Формирование файлов Жюри")
        for str_ in fio_list:
            table_output2_xlsx.getCell("B" + n_rows).setText(str_)
        table_output2_xlsx.getCell("A" + n_rows).setNumber(index)
        table_output2_xlsx.getCell("F" + n_rows).setFormula("=SUM(C" + n_rows + ":D" + n_rows + ":E" + n_rows + ")")
        table_output2_xlsx.getCell("K" + n_rows).setFormula("=SUM(G" + n_rows + ":H" + n_rows + ":I" + n_rows + ":J" + n_rows + ")")
        table_output_xlsx.getCell("W" + n_rows).setFormula("=AVERAGE(AC" + n_rows + ":AE" + n_rows + ")+SUM(" + "X" + n_rows + ":AB" + n_rows + ")")
        table_output_xlsx.getCell("AF" + n_rows).setFormula("=AVERAGE(AG" + n_rows + ":AI" + n_rows + ")")
        table_output_xlsx.getCell("AJ" + n_rows).setFormula("=SUM(AF" + n_rows + ",W" + n_rows + ")")
        document2_xls.saveAs(mydirs_[12].encode('utf-8'))

def main_judges(worker,mydirs_):
    worker.ReportProgress(0, u"Формирование файлов Жюри")
    application = sdk.Application()
    input_file_url_xls = mydirs_[11].encode('utf-8')  # Output с размещением системы рейтингования
    document_xls = application.loadDocument(input_file_url_xls)  # load Сводный
    table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный

    input_file_url2_xls = mydirs_[8].encode('utf-8')  # intput с Жюри1
    document2_xls = application.loadDocument(input_file_url2_xls)  # load Жюри1
    table_output2_xlsx = document2_xls.getBlocks().getTable(0)  # table Жюри1

    n_rows = table_output_xlsx.getRowsCount()  # количество строк
    E4_empty = table_output_xlsx.getCell("E4").getRawValue() == ''

    if not E4_empty:
        cell_range = table_output_xlsx.getCellRange("E4:E" + str(n_rows))
        fio_list = [i.getRawValue() for i in cell_range]

    # Выделение строк в таблице по количеству строк из документов
    rows_count_2 = len(fio_list)
    if rows_count_2 > 1:
        table_output2_xlsx.insertRowAfter(3, copyRowStyle=True, rowsCount=rows_count_2 - 1)
        document2_xls.saveAs(mydirs_[12].encode('utf-8'))
    elif rows_count_2 == 1:
        document2_xls.saveAs(mydirs_[12].encode('utf-8'))

    write_formules(n_rows,fio_list,table_output2_xlsx,table_output_xlsx,document2_xls,mydirs_,worker)
    table_output2_xlsx.getCell("B1").setText("Жюри 2")
    document2_xls.saveAs(mydirs_[13].encode('utf-8'))
    table_output2_xlsx.getCell("B1").setText("Жюри 3")
    document2_xls.saveAs(mydirs_[14].encode('utf-8'))
    worker.ReportProgress(100, u"Завершено формирование файлов Жюри")



