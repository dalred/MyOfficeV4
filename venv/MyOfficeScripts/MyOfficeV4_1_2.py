# -*- coding: utf-8 -*-
from re import compile
import os, time,itertools
from random import choice
from datetime import datetime
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase



def log_file(filename, cell, is_first_error, log):
    if is_first_error:
        log.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " Ошибка в файле: " + " " + filename.encode(
        'utf-8') + " в ячейке " + cell)
    else:
        log[len(log) - 1] = log[len(log)-1]+''.join(", в ячейке " + cell)
    return log


def log_file_rec(log,folderName):
    f = open(folderName+"\\errors.log", "a")
    for i in log:
        f.write(i + "\n")
    f.close()


def clear_content():
    if table_output_xlsx.getRowsCount() < 5 :
        print "Удаление строк не требуется."
        return
    table_output_xlsx.removeRow(4,table_output_xlsx.getRowsCount()-4)
    table_output_xlsx_2.removeRow(4, table_output_xlsx_2.getRowsCount() - 4)
    # rowIndex – индекс строки, начиная с которой удаляются строки;
    # rowsCount – количество удаляемых строк, по умолчанию равно единице.
    # Очистка содержимого ячеек с сохранением форматирования.
    print "Удаление строк."

def extract_txt_doc(filename, path,folderName):
    document_xls_input = application.loadDocument(path)
    table_input = document_xls_input.getBlocks().getTable(0)

    last_name = table_input.getCell('C6').getFormattedValue()
    first_name = table_input.getCell('C8').getFormattedValue()
    middle_name = table_input.getCell('C10').getFormattedValue()

    date_birth = table_input.getCell('C12').getFormattedValue()
    country = table_input.getCell('C14').getFormattedValue()
    district = table_input.getCell('C15').getFormattedValue()
    post_index = table_input.getCell('C16').getFormattedValue()
    region = table_input.getCell('C18').getFormattedValue()
    city = table_input.getCell('C20').getFormattedValue()
    #address = table_input.getCell('C22').getFormattedValue()
    school = table_input.getCell('C22').getFormattedValue()
    school_address = table_input.getCell('C24').getFormattedValue()
    exp = table_input.getCell('C26').getFormattedValue()
    cert_1 = table_input.getCell('C27').getFormattedValue()
    cert_2 = table_input.getCell('C28').getFormattedValue()
    cert_3 = table_input.getCell('C29').getFormattedValue()

    phone = table_input.getCell('C30').getFormattedValue()
    email = table_input.getCell('C32').getFormattedValue()

    parent_fio = table_input.getCell('C34').getFormattedValue()
    parent_email = table_input.getCell('C38').getFormattedValue()
    parent_phone = table_input.getCell('C40').getFormattedValue()
    parent_work = table_input.getCell('C42').getFormattedValue()


    #Раскрашиваем ячейки в анкетах с ошибками
    dict_cells = {
        'C6': last_name,
        'C8': first_name,
        'C10': middle_name,

        'C12': date_birth,
        'C14': country,
        'C15': district,
        'C16': post_index,
        'C18': region,
        'C20': city,
        'C22': school,
        'C24': school_address,
        'C26': exp,
        'C27': cert_1,
        'C28': cert_2,
        'C29': cert_3,

        'C30': phone,
        'C32': email,

        'C34': parent_fio,
        'C38': parent_email,
        'C40': parent_phone,
        'C42': parent_work,
    }
    cell_properties = table_input.getCell('C6').getCellProperties()
    cell_properties.backgroundColor = sdk.ColorRGBA(255, 0, 0, 1)
    is_changed = False
    is_first_error = True
    log = []
    for key, val in dict_cells.iteritems():
        if not val:
            # Помечаем ошибки красным цветом в входном файле
            table_input.getCell(key).setCellProperties(cell_properties)
            log = log_file(filename, key, is_first_error, log)
            is_changed = True
            is_first_error = False
    log_file_rec(log,folderName)

    if is_changed:
        document_xls_input.saveAs((folderName+u"\\Ошибки\\"+ filename).encode('utf-8'))

    full_row_lst = [
                    last_name,
                    first_name,
                    middle_name,
                    date_birth,
                    country,
                    district,
                    post_index,
                    region,
                    city,
                    school,
                    school_address,
                    exp,
                    cert_1 + ", " + cert_2 + ", " +cert_3,
                    phone,
                    email,
                    parent_fio,
                    parent_email,
                    parent_phone,
                    parent_work]
    return full_row_lst



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

def write_table(all_str_sorted_lst,worker):
    print "Запись данных."
    worker.ReportProgress(93, u"Запись данных.")
    current_row = 4
    column = list_xls("W")

    for str_ in all_str_sorted_lst:
        index = str(current_row - 3) # 1
        str_choice = str(choice(['да', 'нет']))
        row_str = str(current_row)  # 4
        table_output_xlsx.getCell("E" + row_str).setText(str_[0] + " " + str_[1] + " " + str_[2])
        table_output_xlsx_2.getCell("B" + row_str).setText(str_[0] + " " + str_[1] + " " + str_[2])
        # A4 set text 1
        k = 1  # Начинаем с B
        #print "str", str_[0]
        for s in str_:
            # print "Столбец ", column, " Строка ", row_str
            # двигаемся построчно
            if k==4 or k==6:
                k += 1
            if k==7:
                k += 1
            table_output_xlsx.getCell(column[k] + row_str).setText(s) # column[1]-B+4 settext из лист str_
            k += 1

        # column = chr(k+66) #Получаем из ASCII, 66 = "B"
        # print "Столбец ", column, " Строка ", row_str
        # table_output_xlsx.getCell(column + row_str).setText(s) #B+4 settext из лист str_

        table_output_xlsx.getCell("AL" + str(current_row)).setContent(str_choice) # Заполнить диапазон X4:кол-во строк.
        for v in str_:
            if not v:
                table_output_xlsx.getCell("AL" + str(current_row)).setText('')
                break
        current_row += 1

def write_formules(worker, number_rows):
    worker.ReportProgress(98, u"Заполнение формул.")
    for i in range(4,number_rows+1):
        index = i - 3
        number_rows=str(i)
        table_output_xlsx.getCell("A" + number_rows).setNumber(index)
        table_output_xlsx_2.getCell("A" + number_rows).setNumber(index)
        table_output_xlsx.getCell("X" + number_rows).setFormula("=AVERAGE(AD" + number_rows + ":AF" + number_rows + ")+SUM(" + "Y" + number_rows + ":AC" + number_rows + ")")
        table_output_xlsx.getCell("AK" + number_rows).setFormula("=SUM(AG" + number_rows + ",X" + number_rows + ")")
        table_output_xlsx.getCell("AG" + number_rows).setFormula("=AVERAGE(AH" + number_rows + ":AJ" + number_rows + ")")
        table_output_xlsx_2.getCell("F" + number_rows).setFormula("=SUM(C" + number_rows + ":D" + number_rows + ":E" + number_rows + ")")
        table_output_xlsx_2.getCell("K" + number_rows).setFormula("=SUM(G" + number_rows + ":H" + number_rows + ":I" + number_rows + ":J" + number_rows + ")")

def error_data(data_error, worker): # Раскрашивает в выходном файле строки с ошибками
    worker.ReportProgress(98, u"Помечаем ошибки")
    for i in data_error:
        cell_properties = sdk.CellProperties()
        cell_properties.backgroundColor = sdk.ColorRGBA(255, 0, 0, 1)
        cell_properties.verticalAlignment = sdk.VerticalAlignment_Center
        cell_range = table_output_xlsx.getCellRange("B" + str(i+3)+":W"+str(i+3))
        cell_range.setCellProperties(cell_properties)

def set_cells_format(number_rows, worker):
    #Покраска зеленым
    print "Применение форматирования."
    worker.ReportProgress(94, u"Применение форматирования.")
    cell_properties = sdk.CellProperties()
    cell_properties.backgroundColor = sdk.ColorRGBA(146, 208, 80, 1)
    #cell_properties.verticalAlignment = sdk.VerticalAlignment_Center
    # Задаем диапозон B4:S - конечная строка
    # Применение форматирования B4:S - конечная строка
    worker.ReportProgress(95, u"Применение форматирования для диапозона B4:W"+ number_rows)
    cell_range = table_output_xlsx.getCellRange("B4:W" + number_rows)
    cell_range.setCellProperties(cell_properties)

    # Задаем вертикальное центрирование
    # для диапозона А4-А - конечная строка.
    worker.ReportProgress(96, u"Применение форматирования для диапозона А4:А" + number_rows)
    cell_properties_aligment = sdk.CellProperties()
    cell_properties_aligment.backgroundColor = sdk.ColorRGBA(0, 0, 0, 0)
    cell_properties_aligment.verticalAlignment = sdk.VerticalAlignment_Center
    cell_range_aligment = table_output_xlsx.getCellRange("A4:A" + number_rows)
    for c in cell_range_aligment:
        c.setCellProperties(cell_properties_aligment)
    #cell_range_aligment.setCellProperties(cell_properties_aligment) Баг-репорт

    # Формат Date для столбца E4
    worker.ReportProgress(97, u"Формат Date для столбца E4")
    cell_range_date = table_output_xlsx.getCellRange("F4:F" + number_rows)
    for c in cell_range_date:
        c.setFormat(sdk.CellFormat_Date)
"""
font_pp = sdk.TextProperties()
font_pp.textColor = sdk.ColorRGBA(0, 0, 0, 1)
font_pp.backgroundColor = sdk.ColorRGBA(255, 255, 0, 1)
pp = sdk.ParagraphProperties()
#pp.alignment = sdk.Alignment_Center
# Установка границ ячеек для всей таблицы
# и горизонтальное форматирование - центр
#

cell_range_borders = table_output_xlsx.getCellRange("A4:AC" + number_rows)
borders_proper = table_output_xlsx.getCell('B3').getBorders()
for c in cell_range_borders:
    c.getRange().setTextProperties(font_pp)
    c.setParagraphProperties(pp)
    c.setBorders(borders_proper) """ # Формат границ




application=None
document_xls=None


def main_(worker, folderName,mydirs_):
    global application
    application = sdk.Application()
    global document_xls
    folder_url = mydirs_[7] #Анкеты
    input_file_url_xls=mydirs_[3].encode('utf-8') #input шаблон с рейтингования
    output_file_url_xls=mydirs_[11].encode('utf-8') #Output с размещением системы рейтингования
    for i in range(12, 15):
        globals()['output_file_url_xls_%d' % (i - 4)] = mydirs_[i].encode('utf-8')
    document_xls = application.loadDocument(input_file_url_xls)
    document_xls_2 = application.loadDocument(mydirs_[8].encode('utf-8'))
    global table_output_xlsx, table_output_xlsx_2
    table_output_xlsx = document_xls.getBlocks().getTable(0)
    table_output_xlsx_2 = document_xls_2.getBlocks().getTable(0)
    all_str_lst = []
    error_index = []
    regex = compile('.*xlsx$')
    for root, dirs, files in os.walk(folder_url):
        del dirs[:] # go only one level deep
        filtered = [i for i in files if regex.match(i)]
        print "Количество файлов: ", len(filtered)
        if not len(filtered):
            raise Exception(u"Нет необходимых xlsx файлов в папке анкеты")
        for filename in filtered:
            percentage = (filtered.index(filename)*91)/len(filtered)
            # print worker.WorkerReportsProgress
            # try:
            worker.ReportProgress(percentage, u"Экспорт анкет.")
            if worker.CancellationPending == True:
                worker.ReportProgress(percentage, u"Отмена задания")
                time.sleep(1)
                return
            # except Exception as e:
            #     print e.message
            path = root + "\\"+filename
            path = path.encode('utf8')
            dict_str = extract_txt_doc(filename,path,folderName)
            all_str_lst.append(dict_str)

    all_str_sorted_lst = sorted(all_str_lst, key=lambda k: k[0])
    error_files_count = 0
    for i in range(0, len(all_str_sorted_lst)):   #от 0 до len записей
        p = len(error_index)
        for v in all_str_sorted_lst[i]:
             if not v:
                 error_index.append(i + 1)
        if len(error_index) > p:
            error_files_count += 1
    clear_content()
    print "Количество записей с ошибками: ", len(error_index)
    print "Количество файлов с ошибками: ", error_files_count
    # Выделение строк в таблице по количеству строк из документов
    n_rows = len(
        all_str_sorted_lst) - table_output_xlsx.getRowsCount() + 3
    n_rows_2 = len(
        all_str_sorted_lst) - table_output_xlsx_2.getRowsCount() + 3
    # необходимое кол-во строк (+3, т.к. 3 строки в заголовке таблицы)
    if n_rows > 0:
        table_output_xlsx.insertRowAfter(3, copyRowStyle=True, rowsCount=n_rows)
        table_output_xlsx_2.insertRowAfter(3, copyRowStyle=True, rowsCount=n_rows_2)

    number_rows = str(table_output_xlsx.getRowsCount())
    print "Количество строк в документе: ", number_rows
    # Записываем результат в таблицу
    write_table(all_str_sorted_lst, worker)
    set_cells_format(number_rows, worker)
    write_formules(worker, table_output_xlsx.getRowsCount())
    error_data(error_index, worker)
    worker.ReportProgress(99, u"Сохранение XLSX.")
    document_xls.saveAs(output_file_url_xls)
    document_xls_2.saveAs(output_file_url_xls_8)
    table_output_xlsx_2.getCell("B1").setText("Жюри 2")
    document_xls_2.saveAs(output_file_url_xls_9)
    table_output_xlsx_2.getCell("B1").setText("Жюри 3")
    document_xls_2.saveAs(output_file_url_xls_10)
    #raise Exception('This is the exception you expect to handle') #Аналог Throw для обработки исключений
    worker.ReportProgress(100, u"Готово.")