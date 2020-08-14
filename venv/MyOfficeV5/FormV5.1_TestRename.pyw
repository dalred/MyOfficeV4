# -*- coding: utf-8 -*-
from re import compile
import os, sys, time, inspect, clr, datetime, tkFileDialog, Tkinter, shutil
from dateutil.relativedelta import relativedelta
from MyOfficeV5_1_TestRename import main_
from Results_v3_2 import main__
from formation_judges import main_judges
from Scores import main_score
clr.AddReference('System')
from System import DateTime as NetDateTime
clr.AddReference('System.Threading.Thread')
from System.Threading import Thread, ThreadStart,ApartmentState
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import *
from System.Drawing import *
from System.ComponentModel import BackgroundWorker
from System.Diagnostics import Process
#https://docs.microsoft.com/ru-ru/dotnet/api/system.componentmodel.backgroundworker?view=netcore-3.1

reload(sys)
sys.setdefaultencoding('utf-8')


formConvert = Form()
progressbar1 = ProgressBar()
combobox1 = ComboBox()
textboxBrowse = TextBox()
start = Button()
canceling = Button()
worker = BackgroundWorker()
datetimepicker1=DateTimePicker()
open = Button()
worker.WorkerReportsProgress = True
worker.WorkerSupportsCancellation = True
textboxBrowse.Text = "D:\Tests\profiles\Тестовая папка".decode('utf-8')
mydir=textboxBrowse.Text
filename = inspect.getframeinfo(inspect.currentframe()).filename
path_dir_root=os.path.dirname(os.path.abspath(filename)).replace("venv\MyOfficeV5", "")
path_dirname_ = os.path.dirname(os.path.abspath(filename)).replace("venv\MyOfficeV5", u"шаблоны")


state_dir = True

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



def do_work(sender, event):
    time.sleep(0.5)
    date_end=datetimepicker1.Value.ToString("dd.MM.yyyy")
    formConvert.Text = str('0%')
    for i in range(3, 11):
        if not (os.path.exists(mydirs_[i])):
            raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i])) #Папку Анкеты тоже проверяем.
    if not (os.path.exists(mydirs_[0])):
        try:
            os.mkdir(mydirs_[0], 0o777)  # Папка Ошибки
            MessageBox.Show(u"Создана папка: " + os.path.basename(mydirs_[0]), u"Информация", MessageBoxButtons.OK,
                        MessageBoxIcon.Information)
        except Exception:
            raise Exception("Неудалось создать папку:"+os.path.basename(mydirs_[0]))
    if combobox1.SelectedIndex == 0:
        if os.path.exists(mydirs_[11]):  # Сводный файл
            dialog_result = MessageBox.Show(u"Вы хотите добавить информацию в: " + os.path.basename(mydirs_[11]),
                                           u"Добавить?", MessageBoxButtons.YesNo,
                                           MessageBoxIcon.Information)
            if dialog_result == DialogResult.Yes:
                pass
            elif dialog_result == DialogResult.No:
                # raise Exception("")
                sender.CancelAsync()
                sender.Dispose()
                return
        else:
            shutil.copyfile(mydirs_[7], mydirs_[11])
        main_(sender, textboxBrowse.Text,mydirs_,str(date_end))
    elif combobox1.SelectedIndex == 1:
        for i in range(8, 12):
            if not (os.path.exists(mydirs_[i])):
                raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i]))
        main_judges(sender, mydirs_)
    elif combobox1.SelectedIndex == 2:
        main__(sender, textboxBrowse.Text,mydirs_)


def Cancel_(sender, event):
    worker.CancelAsync()


def bgWorker_ProgressChanged(sender, event):
    formConvert.Text = str(event.ProgressPercentage) + u"%, " + event.UserState
    progressbar1.Value = event.ProgressPercentage
    if progressbar1.Value==93:
        canceling.Enabled=False


def final(sender,event):
    if event.Error <> None:
        print "Error: ", event.Error.Message
        MessageBox.Show(event.Error.Message,u"Обратитесь к разработчикам!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
    print "RunWorkerCompleted"
    start.Enabled = True
    canceling.Enabled = True
    Application.UseWaitCursor = False
    Cursor.Current = Cursors.Default
    formConvert.Text = 'Задача завершена!'.decode('utf8')
    sender.Dispose()
    time.sleep(1)
    progressbar1.Value = 0
    MessageBox.Show(u"Задача завершена!", u"Информация", MessageBoxButtons.OK,
                    MessageBoxIcon.Information)


def begin_dfile(sender, event):
    start.Enabled = False
    #???foldername = textboxBrowse.Text
    if textboxBrowse.Text == 'folder not specified':
        MessageBox.Show('folder not specified', u"Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    elif combobox1.SelectedIndex == 0 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True
    elif combobox1.SelectedIndex == 1 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True
    elif combobox1.SelectedIndex == 2 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True

def click_open(sender,event):
    if state_dir:
        Process.Start("explorer.exe", textboxBrowse.Text)


def show_dialog(sender, event):
    root = Tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folderName = tkFileDialog.askdirectory(initialdir="/",mustexist=1,title='Пожалуйста укажите корневой каталог: ')
    if folderName:
        open.Enabled = True
        global state_dir
        state_dir = True
        global mydirs_
        textboxBrowse.Text = folderName.replace("/","\\")
        mydirs__ = []
        for i in range(0, 15):
            mydirs__.append((mydirs_[i].replace(mydir,textboxBrowse.Text)))
        mydirs_ = mydirs__ #Без глобал невозможно сделать присвоение статическому полю, так как ты его до этого не объявил.
                            # Использовать можно, но изменять нет.
    else:
        textboxBrowse.Text = 'folder not specified'
        state_dir = False
        open.Enabled = False
    formConvert.Focus()

worker.DoWork += do_work
worker.ProgressChanged += bgWorker_ProgressChanged
worker.RunWorkerCompleted += final



def show_form():
    formConvert.StartPosition = FormStartPosition.CenterScreen
    formConvert.ClientSize = Size(452, 245)
    formConvert.FormBorderStyle = FormBorderStyle.FixedSingle
    formConvert.Name = 'formConvert'
    formConvert.BackColor = SystemColors.ButtonFace
    formConvert.Text = 'Форма для конвертации'.decode('utf8')
    #
    # start
    #
    #
    start.Location = Point(12, 210)
    start.Name = 'start'
    start.Size = Size(110, 30)
    start.TabIndex = 0
    start.Text = 'Start'
    start.Click += begin_dfile
    start.UseCompatibleTextRendering = True
    start.UseVisualStyleBackColor = True
    #
    ## Cancel
    canceling.Location = Point(333, 210)
    canceling.Name = 'canceling'
    canceling.Size = Size(110, 30)
    canceling.TabIndex = 0
    canceling.Text = 'Отмена'.decode('utf-8')
    canceling.UseCompatibleTextRendering = True
    canceling.UseVisualStyleBackColor = True
    canceling.Click += Cancel_
    #
    #open
    open.ImageAlign=ContentAlignment.MiddleCenter
    open.Location = Point(95, 79)
    open.Name = 'open'
    open.Size = Size(32, 27)
    open.TabIndex = 12
    open.UseCompatibleTextRendering = True
    open.UseVisualStyleBackColor = True
    open.Image = Image.FromFile(path_dir_root + "Open-folder-full.png")
    open.Click+=click_open
    #
    #buttonbrowse
    buttonbrowse = Button()
    buttonbrowse.Location = Point(12, 79)
    buttonbrowse.Name = 'buttonbrowse'
    buttonbrowse.Size = Size(77, 27)
    buttonbrowse.TabIndex = 5
    buttonbrowse.Text = u'Обзор'
    buttonbrowse.Click += show_dialog
    buttonbrowse.UseCompatibleTextRendering = True
    buttonbrowse.UseVisualStyleBackColor = True
    #
    # ProgressBar
    #
    #
    progressbar1.Location = Point(12, 170)
    progressbar1.Name = 'progressbar1'
    progressbar1.Size = Size(431, 34)
    progressbar1.Step = 1
    progressbar1.TabIndex = 1
    progressbar1.Value = 0
    progressbar1.ForeColor = Color.Green
    progressbar1.Style = ProgressBarStyle.Continuous

    #
    # combobox1
    #
    combobox1.FormattingEnabled = True
    combobox1.Items.Add(u'1. Обработка Анкет')
    combobox1.Items.Add(u'2.1 Формирование списков Жюри')
    combobox1.Items.Add(u'3. Формирование грамот и писем')
    combobox1.Location = Point(258, 143)
    combobox1.Name = 'combobox1'
    combobox1.Size = Size(185, 21)
    combobox1.TabIndex = 2
    combobox1.SelectedIndex=0
    combobox1.DropDownStyle = ComboBoxStyle.DropDownList
    #combobox1.SelectedIndexChanged += SelectedIndexChanged
    #
    #label Этап
    label = Label()
    label.Location = Point(12, 142)
    label.Name = 'label'
    label.Size = Size(202, 22)
    label.TabIndex = 3
    label.Text = u'Этап обработки'
    label.TextAlign = ContentAlignment.MiddleLeft
    label.UseCompatibleTextRendering = True
    #
    #Путь к размещению файлов
    label1 = Label()
    label1.Location = Point(12, 45)
    label1.Name = 'label1'
    label1.Size = Size(150, 31)
    label1.TabIndex = 4
    label1.Text = u'Путь к корневому каталогу'
    label1.TextAlign = ContentAlignment.MiddleLeft
    label1.UseCompatibleTextRendering = True
    #
    labeldata = Label()
    labeldata.Location = Point(12, 109)
    labeldata.Name = 'labeldata'
    labeldata.Size = Size(185, 20)
    labeldata.TabIndex = 8
    labeldata.Text = u'Дата проведения конкурса:'
    labeldata.TextAlign = ContentAlignment.MiddleLeft
    labeldata.UseCompatibleTextRendering = True
    # TextBox Путь
    textboxBrowse.Location = Point(133, 79)
    textboxBrowse.Name = 'textboxBrowse'
    textboxBrowse.Size = Size(310, 27)
    textboxBrowse.Multiline = True
    textboxBrowse.TabIndex = 6
    textboxBrowse.ReadOnly = True
    #
    #picturebox
    picturebox1=PictureBox()
    picturebox1.Location = Point(368, 4)
    picturebox1.Name = 'picturebox1'
    picturebox1.Size = Size(75, 72)
    picturebox1.SizeMode = PictureBoxSizeMode.CenterImage
    picturebox1.TabIndex = 9
    picturebox1.TabStop = False
    picturebox1.Image = Image.FromFile(path_dir_root+"post_logo.png")
    #
    #
    #datetimepicker1
    datetimepicker1.Format = DateTimePickerFormat.Short
    datetimepicker1.Location = Point(260, 117)
    datetimepicker1.Name = 'datetimepicker1'
    datetimepicker1.Size = Size(183, 20)
    datetimepicker1.TabIndex = 10
    datetimepicker1.Value=NetDateTime(2021, 01, 19)

    #
    # ControlsAdd
    formConvert.BringToFront()
    formConvert.Focus()
    formConvert.Controls.Add(progressbar1)
    formConvert.Controls.Add(canceling)
    formConvert.Controls.Add(start)
    formConvert.Controls.Add(combobox1)
    formConvert.Controls.Add(label)
    formConvert.Controls.Add(label1)
    formConvert.Controls.Add(buttonbrowse)
    formConvert.Controls.Add(textboxBrowse)
    formConvert.Controls.Add(labeldata)
    formConvert.Controls.Add(datetimepicker1)
    formConvert.Controls.Add(picturebox1)
    formConvert.Controls.Add(open)
    Application.Run(formConvert)


t = Thread(ThreadStart(show_form))
t.IsBackground = False
t.ApartmentState = ApartmentState.STA
t.Start()
t.Join()