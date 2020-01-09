#encoding = utf-8
import xlrd
import pandas
import numpy as np
import dateutil


import matplotlib.pyplot as plt
import matplotlib.dates as mdates


#import ttide as tt

import utide
from matplotlib.dates import date2num


import xlsxwriter
import numpy
import datetime
import pycnnum
import lunardate

F = xlsxwriter.Workbook

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False
str_time_from_seconds = lambda seconds: str(round(divmod(seconds, 3600)[0])) + '小时' + str(
    round(divmod(seconds, 3600)[1] / 60)) + '分钟'
str_time_from_second2 = lambda t: '%s小时%s分钟' % (t.seconds // 3600, (t.seconds % 3600) // 60) if (
                                                                                                t.seconds % 3600) % 60 < 30 else  '%s小时%s分钟' % (
t.seconds // 3600, (t.seconds % 3600) // 60 + 1)
constit_tide_level = ['O1', 'Q1', 'P1', 'K1', 'N2', 'M2', 'K2', 'S2', 'M4', 'MS4', 'M6']
pandas.options.display.float_format = '{:,.2f}'.format

def say_out(xxx):
    # 错误提示
    print("*" * 20)
    print(xxx)
    print("*" * 20)

def process(y,threshold):
    # 信号去噪
    rft = np.fft.rfft(y)
    #rft[int(len(rft) / threshold):] = 0
    rft[round(len(rft) * ( 1- threshold / 100)):] = 0
    print('信号去噪结束')
    return np.fft.irfft(rft)

class Tide(object):
    def __init__(self, filename, only_first_sheet = False):
        self.filename = filename
        self.data = {}
        self.year = {}
        self.month = {}
        self.day = {}
        self.html = {}
        self.harmonic_result={}
        self.out = {}
        self.outtxt = {}
        self.sites = xlrd.open_workbook(filename=filename).sheet_names()
        if only_first_sheet:
            self.sites = self.sites[:1]
        for i in self.sites:
            self.data[i] = pandas.read_excel(self.filename, sheet_name=i)
            self.data[i] = self.data[i][['time', 'tide']]


    def tide_statics(self, site=None):
        if site:
            df = self.data[site]
            out = {}
            h = df[df['if_max'] == True]
            l = df[df['if_min'] == True]
            dif = df['diff'].dropna()
            r_t = df.raising_time.dropna().values.mean()
            e_t = df.ebb_time.dropna().values.mean()
            out.update({'mean': df.tide_init.mean()})
            out.update({'h_max': h.tide_init.values.max()})
            out.update({'h_mean': h.tide_init.values.mean()})
            out.update({'l_min': l.tide_init.values.min()})
            out.update({'l_mean': l.tide_init.values.mean()})
            out.update({'dif_mean': dif.mean()})
            out.update({'dif_max': dif.max()})
            out.update({'dif_min': dif.min()})
            out.update({'raise_time': r_t})
            out.update({'ebb_time': e_t})
            self.out.update({site: out})
            self.out_txt(site, out)
        else:
            for i in self.sites:
                self.tide_statics(i)

    def out_txt(self, site, tide_out):
        num_txt = lambda num: str(round(num, 2))
        s = ''
        unit = '米' if (tide_out['dif_max'] < 10) else '厘米'
        s += '潮位观测站' + site + '的平均潮位为' + num_txt(tide_out['mean']) + unit
        s += ',平均高潮位为' + num_txt(tide_out['h_mean']) + unit + ',高潮位最高为' + num_txt(tide_out['h_max']) + unit
        s += ',平均低潮位为' + num_txt(tide_out['l_mean']) + unit + ',低潮位最低为' + num_txt(tide_out['l_min']) + unit
        s += ',平均潮差为' + num_txt(tide_out['dif_mean']) + unit + ',最大潮差为' + num_txt(tide_out['dif_max']) + unit
        s += ',最小潮差为' + num_txt(tide_out['dif_min']) + unit + ',平均涨潮历时为' + str_time_from_second2(tide_out['raise_time'])
        s += ',平均落潮历时为' + str_time_from_second2(tide_out['ebb_time'])
        self.outtxt.update({site: s})

    def preprocess(self,s,threshold=10):
        try:
            self.tide_sheet = pandas.read_excel(self.filename, sheet_name=s)
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
        self.tide_sheet = self.tide_sheet[['time', 'tide']]
        temp_data = self.tide_sheet
        t = []
        if isinstance(temp_data.time[1], pandas.Timestamp) or isinstance(temp_data.time[1],
                                                                               datetime.datetime):
            t = temp_data.time.dt.round('1min')
        else:
            for x in temp_data.time:
                t.append(dateutil.parser.parse(x))
        temp_data['format_time'] = t
        if temp_data.tide.max() - temp_data.tide.min() > 10:
            temp_data['tide_init'] = temp_data.tide / 100
        temp_data = temp_data.dropna(axis=1)
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
        temp_data.tide = process(temp_data.tide,threshold)
        temp_data.index = temp_data['format_time']
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
        self.tide_statics()
        print(s + "站预处理结束")

    def process_one(self, temp_data):

        temp_data['if_min'] = np.r_[True, temp_data.tide.values[1:] < temp_data.tide.values[:-1]] & np.r_[
            temp_data.tide.values[:-1] < temp_data.tide.values[1:], True]
        temp_data['if_max'] = np.r_[True, temp_data.tide.values[1:] > temp_data.tide.values[:-1]] & np.r_[
            temp_data.tide.values[:-1] > temp_data.tide.values[1:], True]

        ##上述方法对一分钟数据不管用2019-3-27尝试用for 循环
        """
        for i in range(3, len(temp_data) - 3):
            if temp_data.iloc[i,1] <  temp_data.iloc[i - 1,1]  and temp_data.iloc[i,1] < temp_data.iloc[i + 1,1]:
                temp_data.iloc[i,- 2 ]  = True
            if temp_data.iloc[i,1] >  temp_data.iloc[i - 1,1]  and temp_data.iloc[i,1] > temp_data.iloc[i + 1,1]:
                temp_data.iloc[i,- 1 ]  = True
        
        """

        temp_data.at[temp_data.index.max(), 'if_min'] = False
        temp_data.at[temp_data.index.max(), 'if_max'] = False
        temp_data.at[temp_data.index.min(), 'if_min'] = False
        temp_data.at[temp_data.index.min(), 'if_max'] = False

        #temp_data.set_value(temp_data.index.max(), 'if_max', False)


        temp_data['diff'] = np.nan

        temp_data['raising_time'] = np.nan
        temp_data['ebb_time'] = np.nan

        def calculate_time_diff(temp_data):
            for x in temp_data.loc[temp_data.if_min == True].format_time:
                try:
                    temp_data.loc[x, 'ebb_time'] = x - temp_data[
                        (temp_data.if_max == True) & (temp_data.format_time < x)].format_time.max()
                    temp_data.loc[x, 'diff'] = temp_data.loc[temp_data[(temp_data.if_max == True) & (
                        temp_data.format_time < x)].format_time.max(), 'tide'] - temp_data.loc[x, 'tide']
                except:
                    pass

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
            return temp_data



                ###############筛选极值
        #####根据潮差与涨落潮历时筛选
        temp_data = calculate_time_diff(temp_data=temp_data)

        temp_data.loc[temp_data[temp_data['diff'] < 1].index, 'if_max'] = False
        temp_data.loc[temp_data[temp_data['diff'] < 1].index, 'if_min'] = False
        temp_data.loc[temp_data[temp_data['diff'] < 1].index, 'raising_time'] = np.nan
        temp_data.loc[temp_data[temp_data['diff'] < 1].index, 'ebb_time'] = np.nan
        temp_data.loc[temp_data[temp_data['diff'] < 1].index, 'diff'] = np.nan

        #2018/11/4 潮差从5改为20

        temp_data.loc[temp_data[temp_data['raising_time'] < datetime.timedelta(hours=2)].index, 'diff'] = np.nan
        temp_data.loc[temp_data[temp_data['raising_time'] < datetime.timedelta(hours=2)].index, 'if_max'] = False
        temp_data.loc[temp_data[temp_data['raising_time'] < datetime.timedelta(hours=2)].index, 'raising_time'] = np.nan

        temp_data.loc[temp_data[temp_data['ebb_time'] < datetime.timedelta(hours=2)].index, 'diff'] = np.nan
        temp_data.loc[temp_data[temp_data['ebb_time'] < datetime.timedelta(hours=2)].index, 'if_min'] = False
        temp_data.loc[temp_data[temp_data['ebb_time'] < datetime.timedelta(hours=2)].index, 'ebb_time'] = np.nan

        ###############根据大小潮个数筛选
        switch = temp_data.loc[temp_data.if_min == True].append(
            temp_data.loc[temp_data.if_max == True]).sort_index().index
        count_filtertime = 0
        count = 0
        count2 = 0
        while (len(temp_data.loc[temp_data.if_min == True]) != len(temp_data.loc[temp_data.if_max == True])) or (
                count > 20 and abs(
                    len(temp_data.loc[temp_data.if_min == True]) - len(temp_data.loc[temp_data.if_max == True])) > 2):
            count += 1
            print('第' + str(count) + '次筛选')
            print('低潮潮位个数' + str(len(temp_data.loc[temp_data.if_min == True])))
            print('高潮潮位个数' + str(len(temp_data.loc[temp_data.if_max == True])))

            if ((len(temp_data.loc[temp_data.if_min == True]) - len(temp_data.loc[temp_data.if_max == True]) == 1) or (
                            len(temp_data.loc[temp_data.if_min == True]) - len(
                            temp_data.loc[temp_data.if_max == True]) == -1)) or count < 30:
                print('极值只多一个')
                count2 += 1
                print('进行第'+str(count2)+'次复查')
                if count2 == 3:
                    print("筛选结束")
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
                                temp_data.loc[switch[i - 1], 'diff'] = np.nan
                                temp_data.loc[switch[i - 1], 'raising_time'] = np.nan
                            else:
                                temp_data.loc[switch[i], 'if_max'] = False
                                temp_data.loc[switch[i], 'diff'] = np.nan
                                temp_data.loc[switch[i], 'raising_time'] = np.nan
                    else:  # 自身低潮
                        if temp_data.loc[switch[i - 1], 'if_min']:  # 后一个也是低潮
                            if temp_data.loc[switch[i], 'tide'] < temp_data.loc[switch[i - 1], 'tide']:  # 前面的i较小,i-1为假
                                temp_data.loc[switch[i - 1], 'if_min'] = False
                                temp_data.loc[switch[i - 1], 'diff'] = np.nan
                                temp_data.loc[switch[i - 1], 'ebb_time'] = np.nan
                            else:
                                temp_data.loc[switch[i], 'if_min'] = False
                                temp_data.loc[switch[i], 'diff'] = np.nan
                                temp_data.loc[switch[i], 'ebb_time'] = np.nan
            switch = temp_data.loc[temp_data.if_min == True].append(
                temp_data.loc[temp_data.if_max == True]).sort_index().index
            count_filtertime += 1
        temp_data = calculate_time_diff(temp_data)

        return temp_data


    def harmonic_analysis_old(self, site, if_init=True):#2019/9/4弃用
        data_h = self.data[site]
        data_h['min'] = data_h.index
        data_h['min']  = data_h['min'].apply(lambda x: x.minute)
        data_h = data_h[data_h['min'] == 0]

        if if_init:
            self.harmonic_result.update({site : tt.t_tide(data_h['tide_init'], dt= 1 ,lat = 30)})
            ###维度采用长江口地区
        else:
            self.harmonic_result.update(
                {site: tt.t_tide(data_h['tide'], dt= 1,lat=30)})

    def harmonic_analysis(self,site,if_init = True,time_zone = 8,lat = 30):#19年9月启用

        ana_data = self.data[site]
        ana_data['UTC_time'] = ana_data['time'] - pandas.Timedelta(time_zone,'h')
        time = date2num(ana_data['UTC_time'].values)
        if if_init:
            result = utide.solve(time, ana_data.tide_init.values, lat=lat, constit=constit_tide_level)
        else:
            result = utide.solve(time, ana_data.tide.values, lat=lat, constit=constit_tide_level)
        self.harmonic_result.update({site: pandas.DataFrame(data = [result['A'],result['A_ci'],result['g'],result['g_ci']],columns = result['diagn']['name'],index = ['振幅','振幅可置信区间','迟角','迟角可置信区间']).to_string()})


    def plot_tide_compare(self, site=None):
        if not site:
            site = self.sites[0]
        temp_data = self.sitedata(site)

        fig, ax = plt.subplots(1, 1, figsize=(100, 100), dpi=400)
        day = mdates.DayLocator()  # every day
        hour = mdates.HourLocator()  # every hour
        dayFmt = mdates.DateFormatter('%m-%d')
        self.tt1 = temp_data.loc[temp_data.if_max == True].format_time  # 高潮时刻
        self.hh1 = temp_data.loc[temp_data.if_max == True].tide  # 高潮潮位
        self.tt2 = temp_data.loc[temp_data.if_min == True].format_time  # 低潮时刻
        self.hh2 = temp_data.loc[temp_data.if_min == True].tide  # 低潮时间

        line1, = ax.plot(temp_data.format_time, temp_data.tide_init, 'o', markersize=1.5, label='原始数据', color='blue')
        line2, = ax.plot(temp_data.format_time, temp_data.tide, linewidth=0.5, color='black', label='去噪以后的数据')

        line3, = ax.plot(self.tt1, self.hh1, marker='*', ls='', markersize=4, color='red', label='高潮时刻')
        line4, = ax.plot(self.tt2, self.hh2, marker='*', ls='', markersize=4, color='green', label='低潮时刻')

        #plt.xlabel('时间', fontsize=18)
        #ax.xaxis.set_major_locator(day)
        #ax.xaxis.set_major_formatter(dayFmt)
        #ax.xaxis.set_minor_locator(hour)
        # ax.legend(handles=[line1, line2],loc = 2,ncol=2, fontsize = 5)
        # ax.legend(handles=[line1, line2], loc=3,bbox_to_anchor = (0,1,0.5,1.1))
        ax.legend(handles=[line1, line2, line3, line4], fontsize=5)
        plt.title(site)
        ax.set_ylabel('潮位')
        self.compare_fig = fig
        return fig
    def sitedata(self, sitename):
        return self.data.get(sitename)

    def sites(self):
        return self.sites

    def out_put_mid_data(self, outfile,sitename=None):
        excel_writer = pandas.ExcelWriter(outfile)
        switch = True
        for s in self.sites:
            if sitename:
                if switch:
                    s = sitename
                    switch = False
                else:
                    continue
            ss =self.data[s]
            ss = ss.drop(['format_time','time'], axis=1)
            ss = ss.replace(np.nan, '', regex=True).replace(False, 'F', regex=True)

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
    def output(self, outfile,sitename=None,time_start=None,time_end = None,reference_system = None):

        chaoshi = "潮时"
        chaogao = "潮高"
        self.file = outfile
        excel_input = xlsxwriter.Workbook(outfile)

        if time_start == pandas.Timestamp(2000,1,1,0,0,0):
            time_start = None
        if time_end == pandas.Timestamp(2000,1,1,0,0,0):
            time_end = None

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
            if sitename :
                if sitename in site:
                    pass
                else:
                    continue
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

                if time_start and time_end :
                    time_start = pandas.Timestamp(time_start)
                    time_end = pandas.Timestamp(time_end)
                    mon = mon.loc[mon.format_time < time_end]
                    mon = mon.loc[mon.format_time >= time_start]
                o_clock = mon.loc[mon.groupby(mon.index.minute == 0).groups[True], ['tide']]

                row1 = count * 42
                count = count + 1  # 第几个月
                ###############################################

                ######################
                for r in range(row1, 3 + row1):
                    sheet.set_row(r, height=20)
                for r in range(row1 + 3, row1 + 41):
                    sheet.set_row(r, height=14.25)

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
                    mm = round(gg.mean(), 1)
                    ss = round(gg.sum(), 1)

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

                sheet.write(row1 + 6 + hangshu, 2, '月 最 高 高 潮 = ' + str(round(alpha * high.tide.max(), 0)),
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

                sheet.write(row1 + 8 + hangshu, 2, '平均涨潮历时:' + str_time_from_seconds(t1),
                            format_cn_a_left1)

                sheet.merge_range(row1 + 8 + hangshu, 21, row1 + 8 + hangshu, 25, '月平均高潮潮高=',
                                  format_cn_a_right1)
                sheet.write(row1 + 8 + hangshu, 26, round(alpha * high.tide.mean()), format_cn_a_left1)
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
                sheet.write(row1 + 6 + hangshu, 9, '月 最 低 低 潮 = ' + str(round(alpha * low.tide.min(), 0)),
                            format_cn_a_left)

                sheet.write(row1 + 7 + hangshu, 9,
                            '潮 时 = ' + str(low.tide.idxmin().month) + '月' + str(low.tide.idxmin().day) + '日' + str(
                                low.tide.idxmin().hour) + '时' + str(low.tide.idxmin().minute)  + '分',
                            format_cn_a_left)

                self.low = mon.loc[mon.groupby(mon.if_min).groups[True], ['tide', 'format_time', 'ebb_time']]
                low = self.low
                for i in low.dropna().index:
                    t2 += (low.loc[i, 'ebb_time']).total_seconds()
                    print('=================================')
                    print(low.loc[i])
                t2 = t2 / len(low.dropna().index)

                print(t2)
                sheet.write(row1 + 8 + hangshu, 9, '平均落潮历时:' + str_time_from_seconds(t2), format_cn_a_left1)
                sheet.merge_range(row1 + 8 + hangshu, 27, row1 + 8 + hangshu, 32, '月平均低潮潮高=',
                                  format_cn_a_right2)
                sheet.merge_range(row1 + 8 + hangshu, 33, row1 + 8 + hangshu, 35, round(alpha * low.tide.mean()),
                                  format_cn_a_left2)
                ###############################################################################
                sheet.write(row1 + 6 + hangshu, 15,
                            '月 平 均 潮 差 = ' + str(round(alpha * numpy.mean(mon['diff'].dropna().values), 0)),
                            format_cn_a_left)
                sheet.write(row1 + 7 + hangshu, 15, '月 最 大 潮 差 = ' + str(round(mon['diff'].max() * alpha, 0)),
                            format_cn_a_left)
                sheet.write(row1 + 8 + hangshu, 15, '月 最 小 潮 差 = ' + str(round(mon['diff'].min() * alpha, 0)),
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
                sheet.write(row1 + 1, 16, '高程系统：'+ reference_system, format_title2)
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
            if han_zi == "一":
                m =  pycnnum.num2cn(lunardate.LunarDate.fromSolarDate(d.year, d.month, d.day).month)
                if len(m) == 1:
                    if m == "一":
                        m = "正"
                    return  m + "月"
                else:
                    m.replace("一十一","冬").replace("一十二","腊").replace("一十",'十') + "月"
            else:
                return '初' + han_zi
        else:
            return han_zi

    def change_altitude(self,sitename,diff):

        for i in [self.month[sitename],self.day[sitename]]:
            for j in i :
                j.tide = j.tide + diff
                j.tide_init = j.tide_init + diff
        self.data[sitename].tide = self.data[sitename].tide +diff
        self.data[sitename].tide_init = self.data[sitename].tide_init + diff

        if not "原始数据" in sitename:
            self.change_altitude(sitename+'原始数据',diff)

    def display(self,site):

        df = self.data[site]
        df2 = df.drop(['time','format_time','if_min','if_max'],1)
        df2.index.names = [None]
        dfs = df2.style.bar(subset=['tide', 'tide_init','diff'], color=['#d65f5f', '#5fba7d'], align='mid')

        h = dfs.render()
        h = h.replace('None',' ').replace('NaT','').replace('0 days','').replace('format_time','时间').replace('tide_init','初始潮位').replace('tide','潮位（去噪）').replace('diff','潮差').replace('raising_time','涨潮时间').replace('ebb_time','落潮时间').replace('nan','').replace('<table id=','<table border="1" id=')
        self.html.update({site:h})



if __name__ == "__main__":
    filename = r"E:\★★★★★项目★★★★★\★★★★埃及汉纳维6000mw清洁煤电项目-水文测验★★★★\实测数据\数据资料--------\二、长期\3、潮位\⑵ 过程资料\10月-11月潮位.xlsx"
    t = Process_Tide(filename,only_first_sheet=True)
    t.preprocess(t.sites[0], 85)
    print(t.harmonic_result)
    print('*****OK*****')
