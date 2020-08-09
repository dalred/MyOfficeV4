# -*- coding: utf-8 -*-
import re
from dateutil.relativedelta import relativedelta
import os, time,itertools,shutil
from datetime import datetime, date
from random import choice
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

def extract_txt_doc(path,folderName,mydirs_,i,lena):
    filename=os.path.basename(path)
    document_xls_input = application.loadDocument(path.encode('utf-8'))
    table_input = document_xls_input.getBlocks().getTable(0)

    last_name = table_input.getCell('C6').getRawValue()
    first_name = table_input.getCell('C8').getRawValue()
    middle_name = table_input.getCell('C10').getRawValue()

    date_birth = table_input.getCell('C12').getFormattedValue()
    country = table_input.getCell('C14').getRawValue()
    district = table_input.getCell('C15').getRawValue()
    post_index = table_input.getCell('C16').getRawValue()
    region = table_input.getCell('C18').getRawValue()
    city = table_input.getCell('C20').getRawValue()
    #address = table_input.getCell('C22').getRawValue()
    school = table_input.getCell('C22').getRawValue()
    school_address = table_input.getCell('C24').getRawValue()
    exp = table_input.getCell('C26').getRawValue()
    cert_1 = table_input.getCell('C27').getRawValue()
    cert_2 = table_input.getCell('C28').getRawValue()
    cert_3 = table_input.getCell('C29').getRawValue()

    phone = table_input.getCell('C30').getRawValue()
    email = table_input.getCell('C32').getRawValue()

    parent_fio = table_input.getCell('C34').getRawValue()
    parent_email = table_input.getCell('C38').getRawValue()
    parent_phone = table_input.getCell('C40').getRawValue()
    parent_work = table_input.getCell('C42').getRawValue()


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
        #document_xls_input.saveAs((mydirs_[0]+"\\"+filename).encode('utf-8'))
        document_xls_input.saveAs((mydirs_[0] + "\\"+"№_"+"_"+last_name+"_"+first_name+"_Ошибка.xlsx").encode('utf-8'))
        os.remove(path)
    else:
        os.rename(path, os.path.dirname(path) + "\\" + "№_"+str(lena + i)+"_"+last_name+"_"+first_name+"_Обработан.xlsx")
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

def write_table(all_str_lst,worker,datetime1_end,n_rows):
    print "Запись данных."
    worker.ReportProgress(93, u"Запись данных.")
    current_row = n_rows
    column = list_xls("W")
    # В разработке =TRUNC(DAYS(G4; F4)/365,242199; 0)
    regex_date = re.compile('^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d$') #mm/dd/year 06/16/1990
    for str_ in all_str_lst:
        index = current_row - 3  # 1
        str_choice = str(choice(['да', 'нет']))
        row_str = str(current_row)  # 4
        if  regex_date.match(str_[3]):
            datetime_birth = convert_time(str_[3])
            time_difference = relativedelta(datetime1_end, datetime_birth)
            difference_in_years = time_difference.years
            table_output_xlsx_11.getCell("H" + row_str).setText(str(difference_in_years))
        else:
            #print "Неверный формат даты, index записи: ", index
            table_output_xlsx_11.getCell("H" + row_str).setText("Error date")
            pass
        table_output_xlsx_11.getCell("E" + row_str).setText(str_[0] + " " + str_[1] + " " + str_[2])
        for i in range(12, 15):
            globals()['table_output_xlsx_%s' % i].getCell("B" + row_str).setText(str_[0] + " " + str_[1] + " " + str_[2])

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
            table_output_xlsx_11.getCell(column[k] + row_str).setText(s) # column[1]-B+4 settext из лист str_
            k += 1

        # column = chr(k+66) #Получаем из ASCII, 66 = "B"
        # print "Столбец ", column, " Строка ", row_str
        # table_output_xlsx.getCell(column + row_str).setText(s) #B+4 settext из лист str_

        table_output_xlsx_11.getCell("AL" + str(current_row)).setContent(str_choice) # Заполнить диапазон X4:кол-во строк.
        for v in str_:
            if not v:
                table_output_xlsx_11.getCell("AL" + str(current_row)).setText('')
                break
        current_row += 1

def write_formules(worker, number_rows,n_rows):
    worker.ReportProgress(98, u"Заполнение формул.")
    current_row = n_rows
    for i in range(current_row,number_rows+1):
        index = i - 3
        number_rows=str(i)
        table_output_xlsx_11.getCell("A" + number_rows).setNumber(index)
        table_output_xlsx_11.getCell("X" + number_rows).setFormula("=AVERAGE(AD" + number_rows + ":AF" + number_rows + ")+SUM(" + "Y" + number_rows + ":AC" + number_rows + ")")
        table_output_xlsx_11.getCell("AK" + number_rows).setFormula("=SUM(AG" + number_rows + ",X" + number_rows + ")")
        table_output_xlsx_11.getCell("AG" + number_rows).setFormula("=AVERAGE(AH" + number_rows + ":AJ" + number_rows + ")")
        for i in range(12, 15):
            globals()['table_output_xlsx_%s' % i].getCell("A" + number_rows).setNumber(index)
            globals()['table_output_xlsx_%s' % i].getCell("F" + number_rows).setFormula("=SUM(C" + number_rows + ":D" + number_rows + ":E" + number_rows + ")")
            globals()['table_output_xlsx_%s' % i].getCell("K" + number_rows).setFormula("=SUM(G" + number_rows + ":H" + number_rows + ":I" + number_rows + ":J" + number_rows + ")")

def error_data(data_error, worker): # Раскрашивает в выходном файле строки с ошибками
    worker.ReportProgress(98, u"Помечаем ошибки")
    for i in data_error:
        cell_properties = sdk.CellProperties()
        cell_properties.backgroundColor = sdk.ColorRGBA(255, 0, 0, 1)
        cell_properties.verticalAlignment = sdk.VerticalAlignment_Center
        cell_range = table_output_xlsx_11.getCellRange("B" + str(i+3)+":W"+str(i+3))
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
    cell_range = table_output_xlsx_11.getCellRange("B4:W" + number_rows)
    cell_range.setCellProperties(cell_properties)

    # Задаем вертикальное центрирование
    # для диапозона А4-А - конечная строка.
    worker.ReportProgress(96, u"Применение форматирования для диапозона А4:А" + number_rows)
    cell_properties_aligment = sdk.CellProperties()
    cell_properties_aligment.backgroundColor = sdk.ColorRGBA(0, 0, 0, 0)
    cell_properties_aligment.verticalAlignment = sdk.VerticalAlignment_Center
    cell_range_aligment = table_output_xlsx_11.getCellRange("A4:A" + number_rows)
    for c in cell_range_aligment:
        c.setCellProperties(cell_properties_aligment)
    #cell_range_aligment.setCellProperties(cell_properties_aligment) Баг-репорт

    # Формат Date для столбца E4
    worker.ReportProgress(97, u"Формат Date для столбца E4")
    cell_range_date = table_output_xlsx_11.getCellRange("F4:F" + number_rows)
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


def convert_time(date_):
    date_str = date_.replace("/", ".").split(".")
    mounth_str = int(date_str[0])
    day_str = int(date_str[1])
    year_str = int(date_str[2])
    datetime_end = date(year_str, mounth_str, day_str)
    return datetime_end

application=None
document_xls=None


def main_(worker, folderName,mydirs_,date_end):
    global application
    application = sdk.Application()
    folder_url = mydirs_[6] #Анкеты
    for i in range(11, 15):
        globals()['output_file_url_xls_%d' % i] = mydirs_[i].encode('utf-8')                                #output path Сводный, Артек 1,2,3
        globals()['document_xls_%s' % i] = application.loadDocument(mydirs_[i].encode('utf-8'))             #load Сводный, Артек 1,2,3
        globals()['table_output_xlsx_%s' % i] = (globals()['document_xls_%s' % i]).getBlocks().getTable(0)  #table Сводный, Артек 1,2,3

    all_str_lst = []
    error_index = []
    lena=0
    regex_done = '№_.*_Обработан\.xlsx'
    regex_not_prep = '(?<!_Обработан).xlsx'  # Попадание в список не обработанных, далее работаем только с ними.
    for root, dirs, files in os.walk(folder_url):
        del dirs[:] # go only one level deep
        filtered_done = [i for i in files if re.search(regex_done, str(i))]
        filtered_not_prep=[i for i in files  if re.search(regex_not_prep,str(i))]
    print "Количество файлов не обработанных: ", len(filtered_not_prep)
    if len(filtered_done+filtered_not_prep)==0:
        raise Exception(u"Нет необходимых xlsx файлов в папке анкеты")
    if len(filtered_not_prep)==0:
        print "Добавление записей не требуется!"
        worker.ReportProgress(100, u"Выполнено.")
        return
    else:
        i = 1
        if len(filtered_done)>0:
            lena = int(filtered_done[len(filtered_done) - 1].split("_")[1])
        for filename in filtered_not_prep:
            percentage = (filtered_not_prep.index(filename)*91)/len(filtered_not_prep)
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
            #path = path.encode('utf8')
            dict_str = extract_txt_doc(path,folderName,mydirs_,i,lena) #folderName - textboxBrowse.Text
            all_str_lst.append(dict_str)
            i += 1

    error_files_count = 0
    for i in range(0, len(all_str_lst)):   #от 0 до len записей
        p = len(error_index)
        for v in all_str_lst[i]:
            if not v:
                #all_str_lst.pop(i)
                error_index.append(i)
        if len(error_index) > p:
            error_files_count += 1
    index_del=set(error_index)
    for index in sorted(index_del, reverse=True):
        del all_str_lst[index]   #Удаление строк из списка в которых есть ошибка
    #clear_content()
    print "Количество записей с ошибками: ", len(error_index)
    print "Количество файлов с ошибками: ", error_files_count
    # Выделение строк в таблице по количеству строк из документов
    n_rows = table_output_xlsx_11.getRowsCount()
    rows_c = len(all_str_lst)
    #print "n_rows,rows_c= ", n_rows ,rows_c
    if rows_c>0:

        if table_output_xlsx_11.getCell("A4").getRawValue()=='':
            #Если пустая строка
            n_rows = n_rows-1
            rows_c=rows_c-1
            for i in range(11, 15):
                globals()['table_output_xlsx_%s' % i].insertRowAfter(n_rows, copyRowStyle=True, rowsCount=rows_c) #добавить в конец количство
        else:
            for i in range(11, 15):
                globals()['table_output_xlsx_%s' % i].insertRowAfter(n_rows-1, copyRowStyle=True, rowsCount=rows_c)  # добавить в конец количство
    number_rows = str(table_output_xlsx_11.getRowsCount())
    print "Количество строк в документе: ", number_rows
    # Записываем результат в таблицу
    #datetime1_end=convert_date(111)
    date_end=convert_time(date_end)
    write_table(all_str_lst, worker,date_end,n_rows+1)
    set_cells_format(number_rows, worker)
    write_formules(worker, table_output_xlsx_11.getRowsCount(),n_rows+1)
    #error_data(error_index, worker)
    worker.ReportProgress(99, u"Сохранение XLSX.")
    for i in range(11, 15):
        globals()['document_xls_%s' % i].saveAs(globals()['output_file_url_xls_%s' % i])
    #raise Exception('This is the exception you expect to handle') #Аналог Throw для обработки исключений
    worker.ReportProgress(100, u"Готово.")