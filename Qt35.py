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


def say_out(xxx):
    # 错误提示
    print(xxx)


def process(y):
    # 信号去噪
    global switch_of_process
    switch_of_process = True
    rft = np.fft.rfft(y)
    rft[int(len(rft) / 5):] = 0
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
                 time2=datetime.datetime(2016, 10, 16, 5, 0)):
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
                        #############################################################
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

            temp_data['raising_time'] = datetime.timedelta(0)
            temp_data['ebb_time'] = datetime.timedelta(0)
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
            if temp_data['diff'].max() <50 :
                temp_data['diff']=temp_data['diff']*100
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

                ###############
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
            self.year[s] = year_tide
            self.month[s] = month_tide
            self.day[s] = day_tide

    def plot_tide_compare(self, site='DongShuiGang',start=20, long=10):
        temp_data = self.sitedata(site)
        self.fig0, self.ax0 = plt.subplots(1, 1, figsize=(100, 10), dpi=300)
        fig = self.fig0
        ax = self.ax0
        day = mdates.DayLocator()  # every day
        hour = mdates.HourLocator()  # every month
        dayFmt = mdates.DateFormatter('%m-%d')
        self.tt1 = temp_data.ix[temp_data.if_max == True].format_time#高潮时刻
        self.hh1 = temp_data.ix[temp_data.if_max == True].tide#高潮潮位
        self.tt2 = temp_data.ix[temp_data.if_min == True].format_time#低潮时刻
        self.hh2 = temp_data.ix[temp_data.if_min == True].tide#低潮时间

        line1, = ax.plot(temp_data.format_time, temp_data.tide_init, 'o', label='Init',color='yellow')
        line2, = ax.plot(temp_data.format_time, temp_data.tide, linewidth=0.5, color='black', label='Processed')

        line3, = ax.plot(self.tt1, self.hh1, marker = '*', ls = '',color='red', label='Highest')
        line4, = ax.plot(self.tt2, self.hh2, marker='*', ls = '',color='magenta', label='Lowest')

        plt.xlim(temp_data.format_time[start], temp_data.format_time[start] + datetime.timedelta(days=long))

        ax.xaxis.set_major_locator(day)
        ax.xaxis.set_major_formatter(dayFmt)
        if (temp_data.time.max() - temp_data.time.min()).days < 60:
            ax.xaxis.set_minor_locator(hour)
        ax.legend(handles=[line1, line2,line3,line4])
        plt.savefig('拟合对比图-' + str(start) + '.png')

    def data2(self):
        return self.data

    def sitedata(self,sitename):
        return  self.data.get(sitename)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example_Widget()
    ex.show()
    ex.showMaximized()
    sys.exit(app.exec_())
    t = Tide(ex.fileName)
    x = Tide('test2.xlsx')
