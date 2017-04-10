# -*- coding: utf-8 -*-
# test
# Form implementation generated from reading ui file 'widget.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!
import sys

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QLabel, QGridLayout
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import sys
import xlrd

import pandas
import numpy as np

import datetime
import dateutil

import matplotlib.pyplot as plt
import matplotlib.dates as mdates

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


def say_out(xxx):
    # 错误提示
    print(xxx)


def process(y):
    # 信号去噪
    global switch_of_process
    switch_of_process = True
    rft = np.fft.rfft(y)
    rft[int(len(rft) / 5):] = 0
    print('信号去噪结束')
    return np.fft.irfft(rft)



class Example_Widget(QWidget):
    def __init__(self):
        super().__init__()

        self.setupUi()

    def setupUi(self):

        self.is_plot = False

        grid = QGridLayout()
        grid.setSpacing(10)
        self.setLayout(grid)

        self.tideButton = QtWidgets.QPushButton('tideButton', self)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(30)
        self.tideButton.setFont(font)
        self.tideButton.setIconSize(QtCore.QSize(20, 16))
        self.tideButton.setObjectName("tideButton")

        grid.addWidget(self.tideButton, 1, 1)

        self.tideButton.clicked.connect(self.open_init_file)

        self.pushButton_2 = QtWidgets.QPushButton(self)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(30)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setIconSize(QtCore.QSize(20, 16))
        self.pushButton_2.setObjectName("pushButton_2")

        grid.addWidget(self.pushButton_2, 2, 1)

        self.pushButton_3 = QtWidgets.QPushButton(self)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(30)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setIconSize(QtCore.QSize(20, 16))
        self.pushButton_3.setObjectName("pushButton_3")

        grid.addWidget(self.pushButton_3, 3, 1)

        self.pushButton_4 = QtWidgets.QPushButton(self)
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(30)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setIconSize(QtCore.QSize(20, 16))
        self.pushButton_4.setObjectName("pushButton_4")

        grid.addWidget(self.pushButton_4, 4, 1)

        self.progressBar = QtWidgets.QProgressBar(self)
        self.progressBar.setGeometry(QtCore.QRect(10, 640, 841, 20))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")

        grid.addWidget(self.progressBar, 1, 10)

        self.dateTimeEdit = QtWidgets.QDateTimeEdit(self)
        self.dateTimeEdit.setGeometry(QtCore.QRect(890, 20, 194, 22))
        self.dateTimeEdit.setObjectName("dateTimeEdit")

        grid.addWidget(self.dateTimeEdit, 1, 10)

        self.dateTimeEdit_2 = QtWidgets.QDateTimeEdit(self)
        self.dateTimeEdit_2.setGeometry(QtCore.QRect(890, 50, 194, 22))
        self.dateTimeEdit_2.setObjectName("dateTimeEdit_2")

        grid.addWidget(self.dateTimeEdit_2, 2, 10)

        self.win_text = QtWidgets.QTextBrowser(self)
        self.win_text.setGeometry(QtCore.QRect(100, 100, 1940, 2200))

        grid.addWidget(self.dateTimeEdit_2, 3, 3)

        self.label = QLabel(self)

        grid.addWidget(self.label, 4, 4)

        self.setWindowTitle("日")

        self.setWindowTitle("Widget")
        self.tideButton.setText("潮位")
        self.pushButton_2.setText("波浪")
        self.pushButton_3.setText("风")
        self.pushButton_4.setText("潮流")

        self.show()

    def change_switch(self, state):
        if state == self.is_plot:
            self.plot()
        else:
            pass

    def show_it(self):
        self.switch_plot = QtWidgets.QCheckBox('是否画对应潮位图', self)
        self.switch_plot.setGeometry(QtCore.QRect(700, 30, 150, 40))

        self.switch_plot.stateChanged.connect(self.change_switch)
        self.switch_plot.toggle()

    def plot(self):
        self.pixmap = QtGui.QPixmap('x.png')
        self.label.setPixmap(self.pixmap)

    def get_time(self):
        self.begin = self.dateTimeEdit.dateTime()
        self.end = self.dateTimeEdit_2.dateTime()

    def open_init_file(self):
        self.show_it()
        self.fileName = QFileDialog.getOpenFileName(self, "选取文件", "E:",
                                                    "Excel文件(*.xlsx);;97-03表格文档 (*.xls)")  # 设置文件扩展名过滤,注意用双分号间隔

        if self.is_plot:
            self.plot()


class Tide(object):
    def __init__(self, filename, time1=datetime.datetime(2000, 1, 1, 00, 0),
                 time2=datetime.datetime(2018, 10, 16, 5, 0)):
        self.data = {}
        self.year = {}
        self.month = {}
        self.day = {}
        for s in xlrd.open_workbook(filename=filename).sheet_names():
            try:
                self.tide_sheet = pandas.read_excel(filename, sheetname=s)
                if 'tide' in self.tide_sheet.columns.values:
                    if 'time' in self.tide_sheet.columns.values:
                        pass
                    else:
                        say_out('Please Check Your Format of the Sheet')
            except:
                say_out('Please Check Your sheets name of the File')

            temp_data = self.tide_sheet

            t = []
            if isinstance(temp_data.time[1], pandas.tslib.Timestamp):
                t = temp_data.time
            else:
                for x in temp_data.time:
                    t.append(dateutil.parser.parse(x))

            temp_data['format_time'] = t
            temp_data['tide_init'] = temp_data.tide
            if len(temp_data)%2  == 1:
                temp_data = temp_data.ix[0:(len(temp_data)-2)]
            temp_data.tide = process(temp_data.tide)
            temp_data.index = temp_data['format_time']

            temp_data = temp_data.ix[temp_data.format_time <= time2]
            temp_data = temp_data.ix[temp_data.format_time >= time1]

            temp_data['if_min'] = np.r_[True, temp_data.tide.values[1:] < temp_data.tide.values[:-1]] & np.r_[
                temp_data.tide.values[:-1] < temp_data.tide.values[1:], True]
            temp_data['if_max'] = np.r_[True, temp_data.tide.values[1:] > temp_data.tide.values[:-1]] & np.r_[
                temp_data.tide.values[:-1] > temp_data.tide.values[1:], True]
            temp_data.set_value(temp_data.index.max(), 'if_min', False)
            temp_data.set_value(temp_data.index.max(), 'if_max', False)
            temp_data.set_value(temp_data.index.min(), 'if_max', False)
            temp_data.set_value(temp_data.index.min(), 'if_min', False)

            temp_data['diff'] = None

            temp_data['raising_time'] = None
            temp_data['ebb_time'] = None
            for x in temp_data.ix[temp_data.if_max == True].format_time:
                try:
                    temp_data.loc[x, 'raising_time'] = x - temp_data[
                        (temp_data.if_min == True) & (temp_data.format_time < x)].format_time.max()
                    temp_data.loc[x, 'diff'] = temp_data.loc[x, 'tide'] - temp_data.loc[temp_data[
                                                                                            (
                                                                                                temp_data.if_min == True) & (
                                                                                                temp_data.format_time < x)].format_time.max(), 'tide']
                except:
                    pass

            for x in temp_data.ix[temp_data.if_min == True].format_time:
                try:
                    temp_data.loc[x, 'ebb_time'] = x - temp_data[
                        (temp_data.if_max == True) & (temp_data.format_time < x)].format_time.max()
                    temp_data.loc[x, 'diff'] = temp_data.loc[temp_data[(temp_data.if_max == True) & (
                        temp_data.format_time < x)].format_time.max(), 'tide'] - temp_data.loc[x, 'tide']
                except:
                    pass
                    ###############筛选极值
            #####根据潮差与涨落潮历时筛选
            if temp_data['diff'].max() < 50:
                temp_data['diff'] = temp_data['diff'] * 100

            temp_data.loc[temp_data[temp_data['diff'] < 10].index, 'diff'] = None
            temp_data.loc[temp_data[temp_data['diff'] < 10].index, 'if_max'] = False
            temp_data.loc[temp_data[temp_data['diff'] < 10].index, 'if_min'] = False
            temp_data.loc[temp_data[temp_data['diff'] < 10].index, 'raising_time'] = 0
            temp_data.loc[temp_data[temp_data['diff'] < 10].index, 'ebb_time'] = 0

            temp_data.loc[temp_data[temp_data['raising_time'] < datetime.timedelta(hours=2)].index, 'diff'] = None
            temp_data.loc[temp_data[temp_data['raising_time'] < datetime.timedelta(hours=2)].index, 'if_max'] = False
            temp_data.loc[temp_data[temp_data['raising_time'] < datetime.timedelta(hours=2)].index, 'raising_time'] = 0

            temp_data.loc[temp_data[temp_data['ebb_time'] < datetime.timedelta(hours=2)].index, 'diff'] = None
            temp_data.loc[temp_data[temp_data['ebb_time'] < datetime.timedelta(hours=2)].index, 'if_min'] = False
            temp_data.loc[temp_data[temp_data['ebb_time'] < datetime.timedelta(hours=2)].index, 'ebb_time'] = 0

            ###############根据大小潮个数筛选
            switch = temp_data.ix[temp_data.if_min == True].append(temp_data.ix[temp_data.if_max == True]).sort_index().index
            count_filtertime = 0
            while (len(temp_data.ix[temp_data.if_min==True]) !=  len(temp_data.ix[temp_data.if_max==True])):

                print('低潮潮位个数'+str(len(temp_data.ix[temp_data.if_min==True]) ))
                print('高潮潮位个数' + str(len(temp_data.ix[temp_data.if_max == True])))
                if ((len(temp_data.ix[temp_data.if_min==True])-len(temp_data.ix[temp_data.if_max==True]) ==1) or (len(temp_data.ix[temp_data.if_min==True])-len(temp_data.ix[temp_data.if_max==True])== -1)):
                    print('极值只多一个，筛选结束')
                    break
                print('进行高低潮筛选————————————————————————————————')


                for i in range(1, len(switch) - 1):
                    ## 时刻i是最大值时
                    if (temp_data.loc[switch[i - 1], 'if_max'] and temp_data.loc[switch[i + 1], 'if_max']) or (
                        temp_data.loc[switch[i - 1], 'if_min'] and temp_data.loc[switch[i + 1], 'if_min']):
                        pass
                    else:
                        if temp_data.loc[switch[i], 'if_max']:  # 自身高潮
                            if temp_data.loc[switch[i - 1], 'if_max']:  # 后一个也是高潮
                                if temp_data.loc[switch[i], 'tide'] > temp_data.loc[switch[i - 1], 'tide']:  # 前面的i大,i-1为假
                                    temp_data.loc[switch[i - 1], 'if_max'] = False
                                    temp_data.loc[switch[i - 1], 'diff'] = None
                                    temp_data.loc[switch[i - 1], 'raising_time'] = None
                                else:
                                    temp_data.loc[switch[i], 'if_max'] = False
                                    temp_data.loc[switch[i ], 'diff'] = None
                                    temp_data.loc[switch[i ], 'raising_time'] = None
                        else:  # 自身低潮
                            if temp_data.loc[switch[i - 1], 'if_min']:  # 后一个也是低潮
                                if temp_data.loc[switch[i], 'tide'] < temp_data.loc[switch[i - 1], 'tide']:  # 前面的i较小,i-1为假
                                    temp_data.loc[switch[i - 1], 'if_min'] = False
                                    temp_data.loc[switch[i - 1], 'diff'] = False
                                    temp_data.loc[switch[i - 1], 'ebb_time'] = False
                                else:
                                    temp_data.loc[switch[i], 'if_min'] = False
                                    temp_data.loc[switch[i], 'diff'] = False
                                    temp_data.loc[switch[i], 'ebb_time'] = False
                switch = temp_data.ix[temp_data.if_min == True].append(
                    temp_data.ix[temp_data.if_max == True]).sort_index().index
                count_filtertime += 1


            year_tide = []
            month_tide = []
            day_tide = []
            for i, j in temp_data.groupby(temp_data.index.year):  #
                year_tide.append(j)
                for ii, jj in j.groupby(j.index.month):
                    month_tide.append(jj)
                    for iii, jjj in jj.groupby(jj.index.day):
                        day_tide.append(jjj)

            self.data[s] = temp_data
            temp_data2 = temp_data
            temp_data2['tide'] = temp_data['tide_init']
            self.data[s+'原始数据'] = temp_data2

            year_tide2 = []
            month_tide2 = []
            day_tide2 = []
            for i, j in temp_data2.groupby(temp_data.index.year):  #
                year_tide2.append(j)
                for ii, jj in j.groupby(j.index.month):
                    month_tide2.append(jj)
                    for iii, jjj in jj.groupby(jj.index.day):
                        day_tide2.append(jjj)

            self.year[s] = year_tide
            self.year[s+'原始数据'] = year_tide2

            self.month[s] = month_tide
            self.month[s+'原始数据'] = month_tide2

            self.day[s] = day_tide
            self.day[s+'原始数据'] = day_tide2

            print(s+"站预处理结束")

    def plot_tide_compare(self, site='DongShuiGang', long=3, date1=None):
        temp_data = self.sitedata(site)
        if date1 == None:
            date1 = temp_data.format_time[1]

        self.fig0, self.ax0 = plt.subplots(1, 1, figsize=(10 * long, 10), dpi=400)
        fig = self.fig0
        ax = self.ax0
        day = mdates.DayLocator()  # every day
        hour = mdates.HourLocator()  # every hour
        dayFmt = mdates.DateFormatter('%m-%d')
        self.tt1 = temp_data.ix[temp_data.if_max == True].format_time  # 高潮时刻
        self.hh1 = temp_data.ix[temp_data.if_max == True].tide  # 高潮潮位
        self.tt2 = temp_data.ix[temp_data.if_min == True].format_time  # 低潮时刻
        self.hh2 = temp_data.ix[temp_data.if_min == True].tide  # 低潮时间

        line1, = ax.plot(temp_data.format_time, temp_data.tide_init, 'o', label='原始数据', color='yellow')
        line2, = ax.plot(temp_data.format_time, temp_data.tide, linewidth=0.5, color='black', label='去噪以后的数据')

        line3, = ax.plot(self.tt1, self.hh1, marker='*', ls='', color='red', label='高潮时刻')
        line4, = ax.plot(self.tt2, self.hh2, marker='*', ls='', color='magenta', label='低潮时刻')

        plt.xlim(date1, date1 + datetime.timedelta(days=long))

        plt.title(site)
        ax.xaxis.set_major_locator(day)
        ax.xaxis.set_major_formatter(dayFmt)
        ax.xaxis.set_minor_locator(hour)
        if (temp_data.time.max() - temp_data.time.min()).days < 60:
            ax.xaxis.set_minor_locator(hour)
        ax.legend(handles=[line1, line2, line3, line4])
        plt.savefig('拟合对比图-' + date1.strftime('%Y-%m-%d') + '.png')

    def ata2(self):
        return self.data

    def sitedata(self, sitename):
        return self.data.get(sitename)


if __name__ == '__main__':
   app = QApplication(sys.argv)
   ex = Example_Widget()
   ex.show()
   ex.showMaximized()
   sys.exit(app.exec_())




