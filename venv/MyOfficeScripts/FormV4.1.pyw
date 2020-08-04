# -*- coding: utf-8 -*-
from re import compile
import os, sys, time, inspect, clr,datetime
from dateutil.relativedelta import relativedelta
from MyOfficeV4_1_2 import main_
from Results_v3_2 import main__
from Scores import main_score
clr.AddReference('System.Threading.Thread')
from System.Threading import Thread, ThreadStart,ApartmentState
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import *
from System.Drawing import *
from System.ComponentModel import BackgroundWorker
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
datetextbox = TextBox()
worker.WorkerReportsProgress = True
worker.WorkerSupportsCancellation = True
textboxBrowse.Text = "D:\Tests\profiles\Тестовая папка".decode('utf-8')
mydirs_=[]
filename = inspect.getframeinfo(inspect.currentframe()).filename
path_dir_root=os.path.dirname(os.path.abspath(filename)).replace("venv\MyOfficeScripts", "")
path_dirname_ = os.path.dirname(os.path.abspath(filename)).replace("\\venv\\MyOfficeScripts", u"\\шаблоны")

def delete_file(dir_):
    """filter=[os.os.path.join(folder_url, f) for f in os.listdir(folder_url) if f <> u"Анкеты"]
    for dir in filter:
        print dir
        time.sleep(0.1)
        remove(dir)"""



def del_(sender, e):
    progressbar1.Value = 0
    delete_file(textboxBrowse.Text)

""" for i in range(0, 4):
        print mydirs_[i]
        delete_file(mydirs_[i])"""

def date_years():
    datetime1 = datetime.date(year_str, mounth_str, day_str)
    datetime2 = datetextbox.Text
    date_str = datetime2.split(".")
    day_str = int(date_str[0])
    mounth_str = int(date_str[1])
    year_str = int(date_str[2])

def do_work(sender, event):
    time.sleep(0.5)
    formConvert.Text = str('0%')
    mydir0 = textboxBrowse.Text + "\\" + u"Ошибки"
    mydir1 = textboxBrowse.Text + "\\" + u"Результаты"
    mydir2 = textboxBrowse.Text + "\\" + u"Грамоты"
    mydir3 = path_dirname_ + "\\" + u"Сводный файл (с учетом новой анкеты участника)_30.07.2020.xlsx"
    mydir4 = path_dirname_ + "\\" + u"Пример Диплома_Имен. падеж_27.07VMyoffice.docx"
    mydir5 = path_dirname_ + "\\" + u"3.1. информац. сообщение детям.docx"
    mydir6 = path_dirname_ + "\\" + u"3.2. информац. сообщение детям.docx"
    mydir7 = textboxBrowse.Text + "\\" + u"Анкеты"
    mydir8= path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_исправленный.xlsx"
    mydir9 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_исправленный.xlsx"
    mydir10 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_исправленный.xlsx"

    mydir11 = textboxBrowse.Text + "\\" + os.path.basename(mydir3)
    mydir12_out = textboxBrowse.Text + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_результаты.xlsx"
    mydir13_out = textboxBrowse.Text + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_результаты.xlsx"
    mydir14_out = textboxBrowse.Text + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_результаты.xlsx"
    mydirs_ = [mydir0, mydir1, mydir2, mydir3, mydir4, mydir5, mydir6, mydir7,mydir8,mydir9,mydir10, mydir11,mydir12_out,mydir13_out,mydir14_out]
    for i in range(3, 11):
        if not (os.path.exists(mydirs_[i])):
            raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i]))
    for i in range(0, 3):
        if not (os.path.exists(mydirs_[i])):
            try:
                os.mkdir(mydirs_[i], 0o777)
                MessageBox.Show(u"Создана папка: " + os.path.basename(mydirs_[i]), u"Информация", MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            except Exception:
                raise Exception("Неудалось создать папку  "+os.path.basename(mydirs_[i]))
    if combobox1.SelectedIndex == 0:
        foldername = textboxBrowse.Text
        for i in range(11, 15):
            if os.path.exists(mydirs_[i]):
                dialogResult = MessageBox.Show(u"Вы хотите перезаписать файл: " + os.path.basename(mydirs_[i]),
                                               u"Перезаписать?", MessageBoxButtons.YesNo,
                                               MessageBoxIcon.Information)
                if dialogResult == DialogResult.Yes:
                    pass
                elif dialogResult == DialogResult.No:
                    # raise Exception("")
                    sender.CancelAsync()
                    sender.Dispose()
                    return
        main_(sender, foldername,mydirs_)
    elif combobox1.SelectedIndex == 1:
        main_score(sender, mydirs_)
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
    state_dir = True
    start.Enabled = False
    #???foldername = textboxBrowse.Text
    if textboxBrowse.Text == 'folder not specified':
        state_dir = False
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



def show_dialog(sender, event):
    folderBrowserDialog1 = FolderBrowserDialog()
    folderBrowserDialog1.RootFolder = 17
    if folderBrowserDialog1.ShowDialog() == 1:
        folderName = folderBrowserDialog1.SelectedPath
        textboxBrowse.Text = folderName
    else:
        textboxBrowse.Text = 'folder not specified'

worker.DoWork += do_work
worker.ProgressChanged += bgWorker_ProgressChanged
worker.RunWorkerCompleted += final



def show_form():
    formConvert.StartPosition = FormStartPosition.CenterScreen
    formConvert.ClientSize = Size(417, 252)
    formConvert.FormBorderStyle = FormBorderStyle.FixedToolWindow
    formConvert.Name = 'formConvert'
    formConvert.Text = 'Форма для конвертации'.decode('utf8')

    #
    # clear
    #
    clear = Button()
    clear.Location = Point(314, 235)
    clear.Name = 'clear'
    clear.Size = Size(91, 40)
    clear.TabIndex = 0
    clear.Click += del_
    clear.Text = 'Очистка'.decode('utf-8')
    clear.UseCompatibleTextRendering = True
    clear.UseVisualStyleBackColor = True
    #
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
    canceling.Location = Point(225, 210)
    canceling.Name = 'canceling'
    canceling.Size = Size(180, 30)
    canceling.TabIndex = 0
    canceling.Text = 'Отмена'.decode('utf-8')
    canceling.UseCompatibleTextRendering = True
    canceling.UseVisualStyleBackColor = True
    canceling.Click += Cancel_
    #
    #
    #
    buttonbrowse = Button()
    buttonbrowse.Location = Point(12, 83)
    buttonbrowse.Name = 'buttonbrowse'
    buttonbrowse.Size = Size(77, 20)
    buttonbrowse.TabIndex = 5
    buttonbrowse.Text = u'Обзор'
    buttonbrowse.Click += show_dialog
    buttonbrowse.UseCompatibleTextRendering = True
    buttonbrowse.UseVisualStyleBackColor = True
    #
    # ProgressBar
    #

    progressbar1.Location = Point(12, 170)
    progressbar1.Name = 'progressbar1'
    progressbar1.Size = Size(393, 34)
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
    combobox1.Items.Add(u'2. Подсчет балов')
    combobox1.Items.Add(u'3. Формирование грамот и писем')
    combobox1.Location = Point(220, 133)
    combobox1.Name = 'combobox1'
    combobox1.Size = Size(185, 21)
    combobox1.TabIndex = 2
    combobox1.SelectedIndex=0
    combobox1.DropDownStyle = ComboBoxStyle.DropDownList
    #combobox1.SelectedIndexChanged += SelectedIndexChanged
    #
    #label Этап
    label = Label()
    label.Location = Point(12, 133)
    label.Name = 'label'
    label.Size = Size(202, 22)
    label.TabIndex = 3
    label.Text = u'Этап обработки'
    label.TextAlign = ContentAlignment.MiddleLeft
    label.UseCompatibleTextRendering = True
    #
    #Путь к размещению файлов
    label1 = Label()
    label1.Location = Point(12, 49)
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
    textboxBrowse.Location = Point(95, 83)
    textboxBrowse.Name = 'textboxBrowse'
    textboxBrowse.Size = Size(310, 20)
    textboxBrowse.TabIndex = 6
    # TextboxDate
    datetextbox.BorderStyle = BorderStyle.FixedSingle
    datetextbox.Location = Point(220, 107)
    datetextbox.Name = 'datetextbox'
    datetextbox.Size = Size(185, 20)
    datetextbox.TabIndex = 7
    datetextbox.TextAlign =  HorizontalAlignment.Left
    datetextbox.Text=datetime.datetime.now().date().strftime('%d.%m.%Y')
    #
    picturebox1=PictureBox()
    picturebox1.Location = Point(330, 5)
    picturebox1.Name = 'picturebox1'
    picturebox1.Size = Size(75, 72)
    picturebox1.SizeMode = PictureBoxSizeMode.CenterImage
    picturebox1.TabIndex = 9
    picturebox1.TabStop = False
    picturebox1.Image = Image.FromFile(path_dir_root+"post_logo.png")

    textboxBrowse.ReadOnly = True
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
    formConvert.Controls.Add(datetextbox)
    formConvert.Controls.Add(picturebox1)
    Application.Run(formConvert)


t = Thread(ThreadStart(show_form))
t.IsBackground = False
t.ApartmentState = ApartmentState.STA
t.Start()
t.Join()