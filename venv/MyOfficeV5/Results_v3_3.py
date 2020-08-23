# -*- coding: utf-8 -*-
import os, sys,re
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk

reload(sys)
sys.setdefaultencoding('utf-8')

def main_results(worker,mydirs_,count):
    regex_count = '^\d+$'
    if not re.search(regex_count, str(count)):
        raise Exception(u"Неправильно указана квота")
    application = sdk.Application()
    document_xls = application.loadDocument(mydirs_[11].encode('utf-8'))
    table_input = document_xls.getBlocks().getTable(0)

    scores_ = []
    n_rows = table_input.getRowsCount()

    def message_lose(table_input, i):
        last_name = table_input.getCell("B" + str(i)).getFormattedValue()
        first_name = table_input.getCell("C" + str(i)).getFormattedValue()
        bookmarks_lose.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
        f_path = (mydirs_[1] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + ' ' + os.path.basename(
            mydirs_[5])).encode('utf-8')
        document_lose.saveAs((f_path))

    def message_win(table_input, i, fio):
        last_name, first_name = fio.split(" ")[0:2]
        bookmarks_win.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
        f_path = (mydirs_[1] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + ' ' + os.path.basename(
            mydirs_[4])).encode('utf-8')
        document_win.saveAs((f_path))

    def paper_win(table_input, i, scores, fio):
        last_name, first_name, middle_name = fio.split(" ")[0:3]
        bookmarks_paper_win.getBookmarkRange('Last_name').replaceText(last_name)
        bookmarks_paper_win.getBookmarkRange('First_middle_name').replaceText(first_name + ' ' + middle_name)
        bookmarks_paper_win.getBookmarkRange('scores').replaceText(str(float("{0:.1f}".format(scores))))

        output_file = mydirs_[2] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + '.pdf'
        document_paper_win.exportAs(str(output_file), sdk.ExportFormat_PDFA1)

    template_win_doc = mydirs_[5].encode('utf-8')
    template_lose_doc = mydirs_[4].encode('utf-8')
    template_paper_win_doc = mydirs_[3].encode('utf-8')
    document_win = application.loadDocument(template_win_doc)
    bookmarks_win = document_win.getBookmarks()
    document_lose = application.loadDocument(template_lose_doc)
    bookmarks_lose = document_lose.getBookmarks()
    document_paper_win = application.loadDocument(template_paper_win_doc)
    bookmarks_paper_win = document_paper_win.getBookmarks()

    for i in range(4, n_rows + 1):
        if float(table_input.getCell("W" + str(i)).getFormattedValue()) > 10:
            fio = table_input.getCell('E' + str(i)).getRawValue()
            balls = float(table_input.getCell('AJ' + str(i)).getRawValue())
            id = int(table_input.getCell('AL' + str(i)).getRawValue())
            full_row_lst = [int(i), balls, id, fio]
            scores_.append(full_row_lst)

    scores_ = sorted(scores_, key=lambda x: (x[1], -x[2]), reverse=True)[0:count]
    sorted_adr = [i[0] for i in scores_]
    k = 0
    for i in range(4, n_rows + 1):
        k += 1
        percentage = int((k * 100) / (n_rows - 3))
        worker.ReportProgress(percentage, u"Формирование грамот и писем.")
        if worker.CancellationPending == True:
            worker.ReportProgress(percentage, u"Отмена задания")
            time.sleep(1)
            return
        if i in sorted_adr:
            fio = scores_[sorted_adr.index(i)][3]
            scores = scores_[sorted_adr.index(i)][1]
            table_input.getCell("AK" + str(i)).setText('Да')
            message_win(table_input, i, fio)
            paper_win(table_input, i, scores, fio)
        else:
            table_input.getCell("AK" + str(i)).setText('Нет')
            message_lose(table_input, i)

    document_xls.saveAs(mydirs_[11].encode('utf8'))