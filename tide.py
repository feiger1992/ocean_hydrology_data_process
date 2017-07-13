#encoding = utf-8
import xlrd
import pandas
import numpy as np
import dateutil

import matplotlib.pyplot as plt
import matplotlib.dates as mdates

import ttide as tt

import xlsxwriter
import numpy
import datetime
import pycnnum
import lunardate

F = xlsxwriter.Workbook

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


def say_out(xxx):
    # 错误提示
    print(xxx)


def process(y):
    # 信号去噪
    rft = np.fft.rfft(y)
    rft[int(len(rft) / 5):] = 0
    print('信号去噪结束')
    return np.fft.irfft(rft)

class Tide(object):
    def __init__(self, filename, only_first_sheet = False,time1=datetime.datetime(2000, 1, 1, 00, 0),
                 time2=datetime.datetime(2018, 10, 16, 5, 0)):
        self.filename = filename
        self.time1 = time1
        self.time2 = time2
        self.data = {}
        self.year = {}
        self.month = {}
        self.day = {}
        self.html = {}
        self.sites = xlrd.open_workbook(filename=filename).sheet_names()
        if only_first_sheet:
            self.sites = self.sites[:1]

    def preprocess(self,s):
        try:
            self.tide_sheet = pandas.read_excel(self.filename, sheetname=s)
            if self.tide_sheet.empty:
                raise ValueError

            if 'tide' in self.tide_sheet.columns.values and 'time' in self.tide_sheet.columns.values:
                pass
            else:
                say_out('Please Check Your Format of the Sheet')
                raise AttributeError
        except:
            say_out('Please Check Your sheets name of the File')
            raise EOFError
        temp_data = self.tide_sheet
        t = []
        if isinstance(temp_data.time[1], pandas.tslib.Timestamp) or isinstance(temp_data.time[1],
                                                                               datetime.datetime):
            t = temp_data.time
        else:
            for x in temp_data.time:
                t.append(dateutil.parser.parse(x))
        temp_data['format_time'] = t
        temp_data['tide_init'] = temp_data.tide / 100
        temp_data = temp_data.dropna()
        self.deltatime = t[1] - t[0]
        if t[0] + (t[1] - t[0]) * (len(t) - 1) != t[len(t) - 1]:
            i_temp = 0
            while i_temp<len(t)-2:
                if t[i_temp+1]-t[i_temp] != self.deltatime:
                    say_out(s+'站时间序列中，'+t[i_temp].strftime('%Y-%m-%d %H:%M')
    +'与'+ t[i_temp+1].strftime('%Y-%m-%d %H:%M')
    +'时间间隔出错,请检查')
                i_temp+=1

        num_time = np.linspace(0, len(t) * self.deltatime.total_seconds() / 3600, len(t))
        temp_data['tide_init'] = temp_data.tide  # 后续调和导出、根据调和分析对数据进行去噪-需要处理###
        if len(temp_data) % 2 == 1:
            temp_data = temp_data.loc[0:(len(temp_data) - 2)]
        temp_data.tide = process(temp_data.tide)
        temp_data.index = temp_data['format_time']
        temp_data = temp_data.loc[temp_data.format_time <= self.time2]
        temp_data = temp_data.loc[temp_data.format_time >= self.time1]
        temp_data = self.process_one(temp_data)
        year_tide = []
        month_tide = []
        day_tide = []
        for i, j in temp_data.groupby(temp_data.index.year):  #
            year_tide.append(j)
            for ii, jj in j.groupby(j.index.month):
                month_tide.append(jj)
                for iii, jjj in jj.groupby(jj.index.day):
                    day_tide.append(jjj)
        x = pandas.concat([temp_data])
        self.data[s + '原始数据'] = temp_data[:]
        self.data[s + '原始数据']['tide'] = self.data[s + '原始数据']['tide_init'][:]
        temp_data2 = self.data[s + '原始数据']
        # self.data.pop(s)
        self.data[s] = x
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
        self.year[s + '原始数据'] = year_tide2
        self.month[s] = month_tide
        self.month[s + '原始数据'] = month_tide2
        self.day[s] = day_tide
        self.day[s + '原始数据'] = day_tide2
        print(s + "站预处理结束")

    def process_one(self, temp_data):

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
        for x in temp_data.loc[temp_data.if_max == True].format_time:
            try:
                temp_data.loc[x, 'raising_time'] = x - temp_data[
                    (temp_data.if_min == True) & (temp_data.format_time < x)].format_time.max()
                temp_data.loc[x, 'diff'] = temp_data.loc[x, 'tide'] - temp_data.loc[temp_data[
                                                                                        (
                                                                                            temp_data.if_min == True) & (
                                                                                            temp_data.format_time < x)].format_time.max(), 'tide']
            except:
                pass

        for x in temp_data.loc[temp_data.if_min == True].format_time:
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
        switch = temp_data.loc[temp_data.if_min == True].append(
            temp_data.loc[temp_data.if_max == True]).sort_index().index
        count_filtertime = 0
        while (len(temp_data.loc[temp_data.if_min == True]) != len(temp_data.loc[temp_data.if_max == True])):

            print('低潮潮位个数' + str(len(temp_data.loc[temp_data.if_min == True])))
            print('高潮潮位个数' + str(len(temp_data.loc[temp_data.if_max == True])))
            if ((len(temp_data.loc[temp_data.if_min == True]) - len(temp_data.loc[temp_data.if_max == True]) == 1) or (
                            len(temp_data.loc[temp_data.if_min == True]) - len(
                            temp_data.loc[temp_data.if_max == True]) == -1)):
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
                                temp_data.loc[switch[i], 'diff'] = None
                                temp_data.loc[switch[i], 'raising_time'] = None
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
            switch = temp_data.loc[temp_data.if_min == True].append(
                temp_data.loc[temp_data.if_max == True]).sort_index().index
            count_filtertime += 1
        return temp_data

    def harmonic_analysis(self, site, if_init=True):
        if if_init:
            return tt.t_tide(self.data[site]['tide_init'], dt=self.deltatime.total_seconds() / 3600)
        else:
            return tt.t_tide(self.data[site]['tide'], dt=self.deltatime.total_seconds() / 3600)

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
        self.tt1 = temp_data.loc[temp_data.if_max == True].format_time  # 高潮时刻
        self.hh1 = temp_data.loc[temp_data.if_max == True].tide  # 高潮潮位
        self.tt2 = temp_data.loc[temp_data.if_min == True].format_time  # 低潮时刻
        self.hh2 = temp_data.loc[temp_data.if_min == True].tide  # 低潮时间

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

    def sitedata(self, sitename):
        return self.data.get(sitename)

    def sites(self):
        return self.sites

    def out_put_mid_data(self, filename):
        excel_writer = pandas.ExcelWriter(filename)
        for s in self.sites:
            ss =self.data[s]
            ss = ss.drop(['format_time','time'],axis=1)
            for i in ss.index:
                t = ss.loc[i, 'raising_time']
                if t:
                    ss.loc[i,'raising_time']=str(t).replace('0 days ','')
            for i in ss.index:
                t = ss.loc[i, 'ebb_time']
                if t:
                    ss.loc[i, 'ebb_time'] = str(t).replace('0 days ', '')
            ss.to_excel(excel_writer,sheet_name=s)
        excel_writer.save()

class Process_Tide(Tide):
    def output(self, outfile):
        chaoshi = "潮时"
        chaogao = "潮高"
        self.file = outfile
        excel_input = xlsxwriter.Workbook(outfile)
        ###########################################基本格式
        format_title1 = excel_input.add_format({
            'bold': 1,
            'font_name': '宋体',
            'font_size': 14,
            'valign': 'vcenter'
        })
        format_title2 = excel_input.add_format({
            'bold': True,
            'font_name': '宋体',
            'font_size': 12,
            'valign': 'vcenter'
        })
        format_cn = excel_input.add_format({
            'border': 1,
            'font_name': '宋体',
            'font_size': 8,
            'valign': 'vcenter',
            'align': 'center'
        })
        format_cn_a_right1 = excel_input.add_format({
            # 月平均高潮潮高
            'left': 1,
            'right': 0,
            'top': 1,
            'bottom': 1,
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'right'
        })
        format_cn_a_right2 = excel_input.add_format({
            # 月平均低潮潮高
            'left': 0,
            'right': 0,
            'top': 1,
            'bottom': 1,
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'right'
        })
        format_cn_a_left = excel_input.add_format({
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'left'
        })
        format_cn_a_left1 = excel_input.add_format({
            # 月平均高潮潮高数字
            # 最底部一行的格式
            'font_name': 'Times New Roman',
            'font_size': 7,
            'bottom': 1,
            'valign': 'vcenter',
            'align': 'left'
        })
        format_cn_a_left2 = excel_input.add_format({
            # 月平均低潮潮高数字
            'font_name': 'Times New Roman',
            'font_size': 7,
            'bottom': 1,
            'right': 1,
            'valign': 'vcenter',
            'align': 'left'
        })
        format_cn_lr = excel_input.add_format({
            'left': 1,
            'right': 1,
            'font_name': '宋体',
            'font_size': 8,
            'valign': 'vcenter',
            'align': 'center'
        })
        format_cn_l = excel_input.add_format({
            'bottom': 1,
            'left': 1,
            'right': 1,
            'font_name': '宋体',
            'font_size': 8,
            'valign': 'vcenter',
            'align': 'left'
        })
        format_cn_r = excel_input.add_format({
            'top': 1,
            'left': 1,
            'right': 1,
            'font_name': '宋体',
            'font_size': 8,
            'valign': 'vcenter',
            'align': 'right'
        })
        format_date = excel_input.add_format({
            'border': 1,
            'num_format': 'mm/dd',
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'center',
        })
        format_time = excel_input.add_format({
            'border': 1,
            'num_format': 'hhmm',
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'center',
        })
        format_num = excel_input.add_format({
            'border': 1,
            'num_format': '0',
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'center',
        })
        format_num_r = excel_input.add_format({
            'border': 1,
            'num_format': '0',
            'font_name': 'Times New Roman',
            'font_color': 'red',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'center',
        })
        format_l = excel_input.add_format({
            'left': 1,
        })
        format_lb = excel_input.add_format({
            'left': 1,
            'bottom': 1
        })

        x = self

        def check(i, t_all):
            # t_all is a list
            if i.date() in t_all:
                return 2
            else:
                t_all.append(i.date())
                return 0

        for site in x.data.keys():
            if x.data[site].tide.max() - x.data[site].tide.min() > 50:
                alpha = 1
            else:
                alpha = 100
            sheet = excel_input.add_worksheet(name=site)
            count = 0
            ############################列宽
            sheet.set_column('A:A', 7)
            sheet.set_column('B:Y', 2.75)
            sheet.set_column('Z:AI', 3.5)
            sheet.set_column('AJ:AJ', 5)
            ##############################
            if max(x.data[site].index) - min(x.data[site].index) < datetime.timedelta(days=50):
                try:
                    x.month[site] = [((x.month[site][0]).append((x.month[site][1]))).sort_index()]
                except:
                    pass

            for mon in x.month[site]:

                row1 = count * 42
                count = count + 1  # 第几个月
                ###############################################

                ######################
                for r in range(row1, 3 + row1):
                    sheet.set_row(r, height=20)
                for r in range(row1 + 3, row1 + 41):
                    sheet.set_row(r, height=14.25)
                o_clock = mon.loc[mon.groupby(mon.index.minute == 0).groups[True], ['tide']]
                ######补全星号
                if len(x.month[site]) != 1:
                    for ii in range(row1 + 6, row1 + 6 + o_clock.index[1].daysinmonth):
                        for jj in range(1, 35):
                            if jj == 33 or jj == 31 or jj == 29 or jj == 27:
                                sheet.write(ii, jj, '**', format_num)
                            else:
                                sheet.write(ii, jj, '*', format_num)
                    hangshu = max(mon.index).daysinmonth
                else:
                    hangshu = (max(mon.index).to_pydatetime().date() - min(
                        mon.index).to_pydatetime().date()).days + 1
                    for ii in range(row1 + 6, row1 + 6 + hangshu):
                        for jj in range(1, 35):
                            if jj == 33 or jj == 31 or jj == 29 or jj == 27:
                                sheet.write(ii, jj, '**', format_num)
                            else:
                                sheet.write(ii, jj, '*', format_num)

                #######################################################表头
                for i in o_clock.index:
                    t = []
                    h = o_clock.get_value(i, 'tide') * alpha
                    if len(x.month[site]) != 1:
                        hanghanghang = i.day
                    else:
                        hanghanghang = (i.to_pydatetime().date() - o_clock.index[0].to_pydatetime().date()).days + 1
                    if len(x.month[site]) != 1:
                        sheet.write(row1 + 5 + i.day, 0, i.date(), format_date)
                        sheet.write(row1 + 5 + i.day, 35, self.solar2lunar_d(i), format_cn)
                        sheet.write(row1 + 5 + i.day, i.hour + 1, h, format_num)
                    else:

                        sheet.write(row1 + 5 + hanghanghang, 0, i.date(), format_date)

                        sheet.write(row1 + 5 + hanghanghang, 0, i.date(), format_date)
                        sheet.write(row1 + 5 + hanghanghang, 35, self.solar2lunar_d(i), format_cn)
                        sheet.write(row1 + 5 + hanghanghang, i.hour + 1, h, format_num)
                    t.append(h)  # 当天的
                    o_clock['day'] = o_clock.index.day
                    gg = o_clock['tide'].groupby(o_clock.day)
                    mm = gg.mean()
                    ss = gg.sum()

                    if len(x.month[site]) != 1:
                        sheet.write(row1 + 5 + i.day, 25, ss.get(i.day) * alpha, format_num)
                        sheet.write(row1 + 5 + i.day, 26, mm.get(i.day) * alpha, format_num)
                        sheet.write(row1 + 6 + i.daysinmonth, 21, '月 合 计 ', format_cn)
                        sheet.write(row1 + 7 + i.daysinmonth, 21, '月 平 均 ', format_cn)  # row1 + 7 + i.daysinmonth, 24,
                    else:
                        sheet.write(row1 + 5 + hanghanghang, 25, ss.get(i.day) * alpha, format_num)
                        sheet.write(row1 + 5 + hanghanghang, 26, mm.get(i.day) * alpha, format_num)
                        sheet.write(row1 + 5 + hangshu + 1, 21, '月 合 计 ', format_cn)
                        sheet.write(row1 + 6 + hangshu + 1, 21, '月 平 均 ', format_cn)
                #if len(x.month[site]) == 1:
                #    hangshu = hangshu + 4  ###############只有一个月时应为2，多月分不知道

                    ########底层画线
                for ii in range(21):
                    sheet.write(row1 + 8 + hangshu, ii, None, format_cn_a_left1)
                ####################################################高低潮位处理###################
                self.high = mon.loc[mon.groupby(mon.if_max).groups[True], ['tide', 'format_time']]
                high = self.high
                self.low = mon.loc[mon.groupby(mon.if_min).groups[True], ['tide', 'format_time']]
                low = self.low
                t1 = 0
                t_count = []

                print('=高潮汇总================================')
                for i in high.dropna().index:
                    if len(x.month[site]) == 1:
                        hanghanghang_1 = (i.to_pydatetime().date() - o_clock.index[0].to_pydatetime().date()).days + 1
                    else:
                        hanghanghang_1 = i.day
                    move1 = check(i, t_count)  # 是否向右移动
                    sheet.write(row1 + 5 + hanghanghang_1, 27 + move1, i, format_time)

                    if i == high.tide.idxmax():
                        sheet.write(row1 + 5 + hanghanghang_1, 28 + move1, high.loc[i, ['tide']] * alpha, format_num_r)
                    else:
                        sheet.write(row1 + 5 + hanghanghang_1, 28 + move1, high.loc[i, ['tide']] * alpha, format_num)

                sheet.write(row1 + 6 + hangshu, 2, '月 最 高 高 潮 = ' + str(int(alpha* high.tide.max())),
                            format_cn_a_left)

                sheet.write(row1 + 7 + hangshu, 2,
                            '潮 时 = ' + str(high.tide.idxmax().month) + '月' + str(high.tide.idxmax().day) + '日' + str(
                                high.tide.idxmax().hour) + '时' + str(high.tide.idxmax().minute) + '分',
                            format_cn_a_left)

                self.high = mon.loc[mon.groupby(mon.if_max).groups[True], ['tide', 'format_time', 'raising_time']]
                high = self.high
                for i in high.dropna().index:
                    t1 += (high.loc[i, 'raising_time']).total_seconds()
                    print(high.loc[i])
                t1 = t1 / len(high.dropna().index)
                # 只有一个月时无法对应#############
                # if len(x.month[site]) == 1:
                #    hangshu += 1
                #   print('+++++')

                sheet.write(row1 + 8 + hangshu, 2, '平均涨潮历时:' + str(int(divmod(t1, 3600)[0])) + '小时' + str(
                    int(divmod(t1, 3600)[1] / 60)) + '分钟',
                            format_cn_a_left1)

                sheet.merge_range(row1 + 8 + hangshu, 21, row1 + 8 + hangshu, 25, '月平均高潮潮高=',
                                  format_cn_a_right1)
                sheet.write(row1 + 8 + hangshu, 26, int(alpha * high.tide.mean()), format_cn_a_left1)
                t2 = 0
                t_count = []
                print('=低潮汇总================================')
                for i in low.dropna().index:
                    if len(x.month[site]) == 1:
                        hanghanghang_2 = (i.to_pydatetime().date() - o_clock.index[0].to_pydatetime().date()).days + 1
                    else:
                        hanghanghang_2 = i.day

                    move1 = check(i, t_count)
                    sheet.write(row1 + 5 + hanghanghang_2, 31 + move1, i, format_time)
                    if i == low.tide.idxmin():
                        sheet.write(row1 + 5 + hanghanghang_2, 32 + move1, low.loc[i, 'tide'] * alpha, format_num_r)
                    else:
                        sheet.write(row1 + 5 + hanghanghang_2, 32 + move1, low.loc[i, 'tide'] * alpha, format_num)
                sheet.write(row1 + 6 + hangshu, 9, '月 最 低 低 潮 = ' + str(int(alpha * low.tide.min())),
                            format_cn_a_left)

                sheet.write(row1 + 7 + hangshu, 9,
                            '潮 时 = ' + str(high.tide.idxmin().month) + '月' + str(high.tide.idxmin().day) + '日' + str(
                                high.tide.idxmin().hour) + '时' + str(high.tide.idxmin().minute) + '分',
                            format_cn_a_left)

                self.low = mon.loc[mon.groupby(mon.if_min).groups[True], ['tide', 'format_time', 'ebb_time']]
                low = self.low
                for i in low.dropna().index:
                    t2 += (low.loc[i, 'ebb_time']).total_seconds()
                    print('=================================')
                    print(low.loc[i])
                t2 = t2 / len(low.dropna().index)

                print(t2)
                sheet.write(row1 + 8 + hangshu, 9, '平均落潮历时:' + str(int(divmod(t2, 3600)[0])) + '小时' + str(
                    int(divmod(t2, 3600 / 60)[1])) + '分钟',
                            format_cn_a_left1)
                sheet.merge_range(row1 + 8 + hangshu, 27, row1 + 8 + hangshu, 32, '月平均低潮潮高=',
                                  format_cn_a_right2)
                sheet.merge_range(row1 + 8 + hangshu, 33, row1 + 8 + hangshu, 35, int(alpha * low.tide.mean()),
                                  format_cn_a_left2)
                ###############################################################################
                sheet.write(row1 + 6 + hangshu, 15,
                            '月 平 均 潮 差 = ' + str(int(alpha * numpy.mean(mon['diff'].dropna().values))),
                            format_cn_a_left)
                sheet.write(row1 + 7 + hangshu, 15, '月 最 大 潮 差 = ' + str(int(mon['diff'].max() * alpha)),
                            format_cn_a_left)
                sheet.write(row1 + 8 + hangshu, 15, '月 最 小 潮 差 = ' + str(int(mon['diff'].min() * alpha)),
                            format_cn_a_left1)
                #############################################################################
                r_1 = row1 + 7
                r_2 = r_1 + hangshu - 1
                sheet.merge_range(row1 + 6 + hangshu, 25, row1 + 6 + hangshu, 26,
                                  '=SUM(Z' + str(r_1) + ':Z' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 25, row1 + 7 + hangshu, 26,
                                  '=AVERAGE(AA' + str(r_1) + ':AA' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 27, row1 + 6 + hangshu, 28,
                                  '=SUM(AC' + str(r_1) + ':AC' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 27, row1 + 7 + hangshu, 28,
                                  '=AVERAGE(AC' + str(r_1) + ':AC' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 29, row1 + 6 + hangshu, 30,
                                  '=SUM(AE' + str(r_1) + ':AE' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 29, row1 + 7 + hangshu, 30,
                                  '=AVERAGE(AE' + str(r_1) + ':AE' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 31, row1 + 6 + hangshu, 32,
                                  '=SUM(AG' + str(r_1) + ':AG' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 31, row1 + 7 + hangshu, 32,
                                  '=AVERAGE(AG' + str(r_1) + ':AG' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 33, row1 + 6 + hangshu, 34,
                                  '=SUM(AI' + str(r_1) + ':AI' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 33, row1 + 7 + hangshu, 34,
                                  '=AVERAGE(AI' + str(r_1) + ':AI' + str(r_2) + ')', format_num)
                ############################
                sheet.write_blank(row1 + 6 + hangshu, 35, None, format_num)
                sheet.write_blank(row1 + 7 + hangshu, 35, None, format_num)
                #######
                sheet.write_blank(row1 + 6 + hangshu, 0, None, format_l)
                sheet.write_blank(row1 + 7 + hangshu, 0, None, format_l)
                sheet.write_blank(row1 + 8 + hangshu, 0, None, format_lb)

                ################################################################################标题
                sheet.write(row1, 0, '潮位观测报表:', format_title1)
                sheet.write(row1 + 1, 0, '工程名称：', format_title2)
                sheet.write(row1 + 1, 0, '工程名称：', format_title2)
                sheet.write(row1 + 2, 0, '海区：', format_title2)
                sheet.write(row1 + 1, 16, '高程系统：原始数据高程', format_title2)
                sheet.write(row1 + 2, 16,
                            '观测日期：' + mon.format_time.min().strftime('%Y/%m/%d') + '至' + mon.format_time.max().strftime(
                                '%Y/%m/%d'), format_title2)
                sheet.write(row1 + 1, 29, '测站：' + str(site).replace('原始数据', ''), format_title2)
                sheet.write(row1 + 2, 29, '单位：cm', format_title2)
                sheet.write(row1 + 3, 0, chaoshi, format_cn_r)
                sheet.write(row1 + 4, 0, chaogao, format_cn_lr)
                sheet.write(row1 + 5, 0, '日期', format_cn_l)
                for c in range(1, 25):
                    sheet.merge_range(row1 + 3, c, row1 + 5, c, c - 1, format_cn)

                sheet.merge_range(row1 + 3, 25, row1 + 5, 25, '合计', format_cn)
                sheet.merge_range(row1 + 3, 26, row1 + 5, 26, '平均', format_cn)
                sheet.merge_range(row1 + 3, 27, row1 + 3, 28, '高潮', format_cn)
                sheet.write(row1 + 4, 27, chaoshi, format_cn)
                sheet.write(row1 + 5, 27, '', format_cn)
                sheet.write(row1 + 4, 28, chaogao, format_cn)
                sheet.write(row1 + 5, 28, '', format_cn)
                sheet.merge_range(row1 + 3, 29, row1 + 3, 30, '高潮', format_cn)
                sheet.write(row1 + 4, 29, chaoshi, format_cn)
                sheet.write(row1 + 5, 29, '', format_cn)
                sheet.write(row1 + 4, 30, chaogao, format_cn)
                sheet.write(row1 + 5, 30, '', format_cn)
                sheet.merge_range(row1 + 3, 31, row1 + 3, 32, '低潮', format_cn)
                sheet.write(row1 + 4, 31, chaoshi, format_cn)
                sheet.write(row1 + 5, 31, '', format_cn)
                sheet.write(row1 + 4, 32, chaogao, format_cn)
                sheet.write(row1 + 5, 32, '', format_cn)
                sheet.merge_range(row1 + 3, 33, row1 + 3, 34, '低潮', format_cn)
                sheet.write(row1 + 4, 33, chaoshi, format_cn)
                sheet.write(row1 + 5, 33, '', format_cn)
                sheet.write(row1 + 4, 34, chaogao, format_cn)
                sheet.write(row1 + 5, 34, '', format_cn)
                sheet.merge_range(row1 + 3, 35, row1 + 5, 35, '农历', format_cn)

                #########################################底部划线
                print('**********************************************')
                print('正在写入' + site + '站第' + str(count) + '个月')
            print(str(site) + '写入结束')

        print(self.file + '文件写入完成')
        excel_input.close()

    def solar2lunar_d(self, d):
        # 将阳历转化为阴历
        # 汉字格式改写
        han_zi = pycnnum.num2cn(lunardate.LunarDate.fromSolarDate(d.year, d.month, d.day).day)
        han_zi = han_zi.replace('一十', '十')
        if len(han_zi) == 3:
            return han_zi.replace('二十', '廿')
        if len(han_zi) == 1:
            return '初' + han_zi
        else:
            return han_zi

    def outfile(self):
        return self.file

    def high(self):
        return self.high

    def display(self,site=None):

        if len(self.sites)==1:
            site = self.sites[0]
        df = self.data[site]
        df2 = df.drop(['time','format_time','if_min','if_max'],1)
        df2.index.names = [None]
        dfs = df2.style.bar(subset=['tide', 'tide_init','diff'], color=['#d65f5f', '#5fba7d'], align='mid')

        h = dfs.render()
        h = h.replace('None',' ').replace('NaT','').replace('0 days','').replace('format_time','时间').replace('tide_init','初始潮位').replace('tide','潮位（去噪）').replace('diff','潮差').replace('raising_time','涨潮时间').replace('ebb_time','落潮时间').replace('nan','').replace('<table id=','<table border="1" id=')
        self.html.update({site:h})

if __name__ == "__main__":
    for i in [11]:
        filename1 = r"E:\★★★★★CODE★★★★★\Git\codes\ocean_hydrology_data_process\测试潮位 - 副本.xls"
        filename2 = filename1.replace(filename1.split('.')[-1],'中间数据.xlsx')
        filename3 = filename1.replace(filename1.split('.')[-1],'结果.xlsx')
        t = Process_Tide(filename1)
        for i in t.sites:
            t.preprocess(i)
        """t.display()
        with open(r"E:\★★★★★CODE★★★★★\程序调试对比\潮汐模块\对比潮位特征值（东水港村）\中间数据2012-3.html", 'w') as f:
            for _,v in t.html.items():
                f.write(v)
        excel_writer = pandas.ExcelWriter(filename2)
        for key,v in t.data.items():
            if '原始数据' not in key:
                v.to_excel(excel_writer,sheet_name=key)
                excel_writer.save()
        t.output(filename3)"""
        print('*****OK*****')


