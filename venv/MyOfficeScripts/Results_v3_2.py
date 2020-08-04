# -*- coding: utf-8 -*-
import os, sys
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk

reload(sys)
sys.setdefaultencoding('utf-8')
def main__(worker, folderName,mydirs_):
    def message_lose(table, i):
        last_name = table.getCell("B" + str(i)).getFormattedValue()
        first_name = table.getCell("C" + str(i)).getFormattedValue()
        bookmarks_lose.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
        f_path = (mydirs_[1]+"\\"+'№'+ str(i - 3) + ' '+ last_name + ' ' + first_name+ ' ' +os.path.basename(mydirs_[5])).encode('utf-8')
        document_lose.saveAs((f_path))
    def message_win(table, i):
        last_name = table.getCell("B" + str(i)).getFormattedValue()
        first_name = table.getCell("C" + str(i)).getFormattedValue()
        bookmarks_win.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
        f_path = (mydirs_[1]+"\\"+'№'+ str(i - 3) + ' '+ last_name + ' ' + first_name+ ' ' +os.path.basename(mydirs_[6])).encode('utf-8')
        document_win.saveAs((f_path))
    def paper_win(table, i):
        last_name = table.getCell("B" + str(i)).getFormattedValue()
        first_name = table.getCell("C" + str(i)).getFormattedValue()
        middle_name = table.getCell("D" + str(i)).getFormattedValue()
        bookmarks_paper_win.getBookmarkRange('Last_name').replaceText(last_name)
        bookmarks_paper_win.getBookmarkRange('First_middle_name').replaceText(first_name + ' ' + middle_name)
        bookmarks_paper_win.getBookmarkRange('scores').replaceText('100')

        output_file =mydirs_[2]+"\\"+'№' +str(i - 3) + ' ' + last_name + ' ' + first_name + '.pdf'
        document_paper_win.exportAs(str(output_file), sdk.ExportFormat_PDFA1)


    input_xls = (folderName + "\\"+os.path.basename(mydirs_[3])).encode('utf-8')
    template_win_doc = mydirs_[6].encode('utf-8')
    template_lose_doc = mydirs_[5].encode('utf-8')
    template_paper_win_doc = mydirs_[4].encode('utf-8')
    application = sdk.Application()
    document_win = application.loadDocument(template_win_doc)
    bookmarks_win = document_win.getBookmarks()
    document_lose = application.loadDocument(template_lose_doc)
    bookmarks_lose = document_lose.getBookmarks()
    document_paper_win = application.loadDocument(template_paper_win_doc)
    bookmarks_paper_win = document_paper_win.getBookmarks()
    try:
        document_xls_intput = application.loadDocument(input_xls)
    except Exception:
        raise Exception('This is the exception you expect to handle')
    table = document_xls_intput.getBlocks().getTable(0)
    last_row = table.getRowsCount()
    k = 0
    for i in range(4, last_row + 1):
        k += 1
        percentage = int((k * 100) / (last_row - 3))
        worker.ReportProgress(percentage, u"Формирование грамот и писем.")
        if worker.CancellationPending == True:
            worker.ReportProgress(percentage, u"Отмена задания")
            time.sleep(1)
            return
        if table.getCell("AC" + str(i)).getFormattedValue() == 'да':
            message_win(table, i)
            paper_win(table, i)
        elif table.getCell("AC" + str(i)).getFormattedValue() == 'нет':
            message_lose(table, i)