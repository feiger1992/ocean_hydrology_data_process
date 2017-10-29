from functools import reduce
from dxfwrite import DXFEngine as dxf

import matplotlib.pyplot as plt
import operator
import pandas
import numpy as np
import datetime

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

east = lambda v, d: v * np.sin(d / 180 * np.pi)
north = lambda v, d: v * np.cos(d / 180 * np.pi)
velocity = lambda v_e, v_n: np.sqrt(v_e ** 2 + v_n ** 2)
direction = lambda v_e, v_n: 180 / np.pi * np.arctan2(v_e, v_n) if (v_e > 0)else   180 / np.pi * np.arctan2(v_e,
                                                                                                            v_n) + 360
small_diff_dir = lambda d1, d2: (d1 - d2 if d1 - d2 < 180 else d2 + 360 - d1) if d1 > d2 else  (
d2 - d1 if d2 - d1 < 180 else d1 + 360 - d2)
is_d1_in_d2_and_d3 = lambda d1, d2, d3: True if (
small_diff_dir(d1, d3) + small_diff_dir(d1, d2) == small_diff_dir(d2, d3)) else False
dir_in_360b = lambda d: (d - 360 if d > 360 else d) if d > 0 else  360 + d
dir_in_360 = lambda d: dir_in_360b(d) if (dir_in_360b(d) > 0 and (dir_in_360b(d) < 360)) else dir_in_360b(
    dir_in_360b(d))

str_from_datatime64 = lambda t: str(t)[:-31][-19:].replace('T', ' ')
str_from_float = lambda t, n=2: str(round(t, n))

def aoe(fun, v_es, v_ns):
    vs = []
    for i, j in zip(v_es, v_ns):
        v = fun(i, j)
        vs.append(v)
    return vs

def is_time_of1(d):
    def raise_or_not(angle1, ang=0):  # 涨/落潮主方向
        if angle1 > 180:
            raise ValueError
            return "Wrong with the angle"
        if angle1 < 90:
            angle1 = angle1 + 90
        else:
            angle1 = angle1 - 90
        angle2 = angle1 + 180
        if min(abs(d - angle1), abs(d - angle2), 360 - abs(d - angle2), 360 - abs(d - angle1)) < ang:
            return 0  # 转流
        elif angle1 < d < angle2:
            return 1  # 涨/落潮
        else:
            return -1  # 落/涨潮

    return raise_or_not

def is_time_of(d):
    def raise_or_not(angle, ang=0):  # angle 为落潮方向
        if ang > 90:
            raise ValueError
            return "Wrong with the angle"
        d1 = dir_in_360(angle + 90)
        d2 = dir_in_360(angle - 90)
        if min(abs(d - d1), abs(d - d2), 360 - abs(d - d1), 360 - abs(d - d2)) < ang:
            return 0  # 转流
        elif small_diff_dir(d, angle) < 90:  # 同向
            return 1
        else:
            return -1

    return raise_or_not

def get_duration(time_v_d):  # 统计涨落潮时间
    time_v_d = time_v_d[time_v_d.timeof != 0]
    time_v_d.index = range(0, len(time_v_d))
    raise_time = 0  # 此时无法区分涨落，仅作标识用途
    ebb_time = 0
    raise_duration = datetime.timedelta(0)
    ebb_duration = datetime.timedelta(0)
    raise_last = False
    ebb_last = False
    raise_times = []
    ebb_times = []
    for i in range(1, len(time_v_d)):
        if (time_v_d.loc[i, 'timeof'] == 1) and (time_v_d.loc[i - 1, 'timeof'] == 1) and raise_last:  # 涨潮last
            raise_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time'])
            continue
        if (time_v_d.loc[i, 'timeof'] == -1) and (time_v_d.loc[i - 1, 'timeof'] == -1) and ebb_last:  # 落潮last
            ebb_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time'])
            continue
        if (time_v_d.loc[i, 'timeof'] == -1) and (time_v_d.loc[i - 1, 'timeof'] == 1):  # 涨潮end,直接变落潮
            if raise_last:
                raise_time += 0.5
                raise_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time']) / 2
                raise_last = False
                raise_times.append(raise_duration)
                raise_duration = datetime.timedelta(0)
            ebb_time += 0.5
            ebb_last = True
            ebb_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time']) / 2
            continue
        if (time_v_d.loc[i, 'timeof'] == 1) and (time_v_d.loc[i - 1, 'timeof'] == -1):  # 落潮end,直接变涨潮
            if ebb_last:
                ebb_time += 0.5
                ebb_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time']) / 2
                ebb_last = False
                ebb_times.append(ebb_duration)
                ebb_duration = datetime.timedelta(0)
            raise_time += 0.5
            raise_last = True
            raise_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time']) / 2
            continue
    # print('raise_time=' + str(raise_time))  # 次数
    # 第一个时间为首次出现的方向对应次数  # 次数
    return {'durations_1': raise_times, 'durations_2': ebb_times, 'times1': raise_time, 'times2': ebb_time,
            'last_time1': reduce(operator.add, raise_times) / int(raise_time),
            'last_time2': reduce(operator.add, ebb_times) / int(ebb_time)}


class Current_pre_process(object):
    def __init__(self, filename, sheet_name=None, is_VD=True):
        if 'csv' in filename:
            self.df = pandas.read_csv(filename)
        if 'xlsx' in filename:
            f = pandas.ExcelFile(filename)
            if not sheet_name:
                sheet_name = f.sheet_names[0]
            self.name = sheet_name
            self.df = f.parse(sheetname=sheet_name)
        self.df = self.df.replace(to_replace=3276.7, value=np.nan)
        self.df = self.df.drop([i for i in self.df if 'Unnamed' in i], axis=1)

        if 't' in self.df.columns:
            if 'time' in self.df.columns:
                raise ValueError('文件中没有时间数据')
        try:
            self.df['t'] = pandas.to_datetime(self.df['t'])
        except:
            pass
        self.t = self.df['t']

        if is_VD:
            d_i = [d for d in self.df.columns if 'd' in d]
            v_i = [v for v in self.df.columns if 'v' in v]
            if not (d_i and v_i):
                raise ValueError('文件中没有流速流向数据')

            self.d = self.df[d_i]
            self.v = self.df[v_i]

            self.l = self.get_less_index(self.d, self.v)
            self.e, self.n = self.convert_to_e_n()

        else:
            e_i = [e for e in self.df.columns if 'e' in e]
            n_i = [n for n in self.df.columns if 'n' in n]
            if not (e_i and n_i):
                raise ValueError('文件中缺少分向流速数据')
            self.n = self.df[n_i]
            self.e = self.df[e_i]
            self.l = self.get_less_index(self.n, self.e)
            self.v, self.d = self.convert_to_v_d()

    def convert_to_e_n(self):
        e = pandas.DataFrame()
        n = pandas.DataFrame()
        for i in range(1, len(self.d.columns) + 1):
            try:
                e['e' + str(i)] = east(self.v['v' + str(i)], self.d['d' + str(i)])
                n['n' + str(i)] = north(self.v['v' + str(i)], self.d['d' + str(i)])
            except:
                e['e' + str(i)] = east(self.v['v' + str(i)].convert_objects(convert_numeric=True),
                                       self.d['d' + str(i)].convert_objects(convert_numeric=True))
                n['n' + str(i)] = north(self.v['v' + str(i)].convert_objects(convert_numeric=True),
                                        self.d['d' + str(i)].convert_objects(convert_numeric=True))
        return e, n

    def convert_to_v_d(self):
        v = pandas.DataFrame()
        d = pandas.DataFrame()
        for i in range(1, len(self.e.columns) + 1):
            try:
                v['v' + str(i)] = aoe(velocity, self.e['e' + str(i)], self.n['n' + str(i)])
                d['n' + str(i)] = aoe(direction, self.e['e' + str(i)], self.n['n' + str(i)])
            except:
                v['e' + str(i)] = velocity(self.e['e' + str(i)].convert_objects(convert_numeric=True),
                                           self.n['n' + str(i)].convert_objects(convert_numeric=True))
                d['e' + str(i)] = direction(self.e['e' + str(i)].convert_objects(convert_numeric=True),
                                            self.n['n' + str(i)].convert_objects(convert_numeric=True))
        return v, d

    def convert_to_EN_file(self, outfile):
        if not 'csv' in outfile:
            raise ValueError("输出文件为csv文件")
        self.merge_save_csv(self.e, self.n, outfile)

    def convert_to_VD_file(self, outfile):
        if not 'csv' in outfile:
            raise ValueError("输出文件为csv文件")
        self.merge_save_csv(self.v, self.d, outfile)

    def merge_save_csv(self, v1, v2, outfile):
        x = pandas.merge(v1, v2, how='outer', on=None, left_index=True, right_index=True)
        x = pandas.merge(pandas.DataFrame(self.t), x, how='outer', left_index=True, right_index=True)
        x.to_csv(outfile, sep=',', index=False, encoding='utf-8')

    def get_less_index(slef, df_1, df_2):
        def find_NaN_index(df):
            l = []
            for i in range(len(df)):
                tt = df.iloc[i].convert_objects(convert_numeric=True)
                s = list(tt.apply(np.isnan))
                if True in s:
                    l.append(s.index(True) - 1)
                else:
                    l.append(len(s) - 1)
            return l  # l的层数从0开始算

        l1 = find_NaN_index(df_1)
        l2 = find_NaN_index(df_2)
        if len(l1) != len(l2):
            raise ValueError
        l = []
        for i, j in zip(l1, l2):
            if i < j:
                l.append(i)
            else:
                l.append(j)
        return l

    def fen_ceng(self, zongcengshu, t, bin, first_bin, top_ratio, button_ratio):
        depth = first_bin + bin * (zongcengshu + 1)
        top_e = self.e['e1'][t] * top_ratio
        top_n = self.n['n1'][t] * top_ratio
        button_e = self.e['e' + str(zongcengshu)][t] * button_ratio
        button_n = self.n['n' + str(zongcengshu)][t] * button_ratio
        h_8 = depth * 0.8 - first_bin
        h_2 = depth * 0.2 - first_bin
        h_4 = depth * 0.4 - first_bin
        h_6 = depth * 0.6 - first_bin

        def quan_zhong(x):
            x = round(x, 3)
            x2 = round(x)
            if x < 1:
                return 1, 1, 0.5, 0.5
            if x2 > x:
                return x2 - 1, x2, x - (x2 - 1), x2 - x
            else:
                return x2, x2 + 1, x - x2, x2 + 1 - x  # 返回的x为第x层，从1算起

        def jia_quan(df_e, df_n, t, i, j, i_r, j_r):
            return df_e['e' + str(i)][t] * i_r + df_e['e' + str(j)][t] * j_r, df_n['n' + str(i)][t] * i_r + \
                   df_n['n' + str(j)][t] * j_r

        i, j, i_r, j_r = quan_zhong(h_2)
        h_2_e, h_2_n = jia_quan(self.e, self.n, t, i, j, i_r, j_r)

        i, j, i_r, j_r = quan_zhong(h_4)
        h_4_e, h_4_n = jia_quan(self.e, self.n, t, i, j, i_r, j_r)

        i, j, i_r, j_r = quan_zhong(h_6)
        h_6_e, h_6_n = jia_quan(self.e, self.n, t, i, j, i_r, j_r)

        i, j, i_r, j_r = quan_zhong(h_8)
        h_8_e, h_8_n = jia_quan(self.e, self.n, t, i, j, i_r, j_r)

        ave_e = (top_e + 2 * h_2_e + 2 * h_4_e + 2 * h_6_e + 2 * h_8_e + button_e) / 10
        ave_n = (top_n + 2 * h_2_n + 2 * h_4_n + 2 * h_6_n + 2 * h_8_n + button_n) / 10
        return top_e, top_n, h_2_e, h_2_n, h_4_e, h_4_n, h_6_e, h_6_n, h_8_e, h_8_n, button_e, button_n, ave_e, ave_n, depth

    def all_fenceng(self, bin=1, first_bin=0.7, top_ratio=1.05, button_ratio=0.95):
        columns = ['top', 'h_2', 'h_4', 'h_6', 'h_8', 'button', 'ave']

        self.ee = pandas.DataFrame(columns=columns)
        self.nn = pandas.DataFrame(columns=columns)
        self.vv = pandas.DataFrame(columns=columns)
        self.dd = pandas.DataFrame(columns=columns)
        all_depth = []
        for i in self.df.index:
            top_e, top_n, h_2_e, h_2_n, h_4_e, h_4_n, h_6_e, h_6_n, h_8_e, h_8_n, button_e, button_n, ave_e, ave_n, depth = self.fen_ceng(
                self.l[i], i, bin, first_bin, top_ratio, button_ratio)
            e = pandas.DataFrame([[top_e, h_2_e, h_4_e, h_6_e, h_8_n, button_e, ave_e]], columns=columns)
            n = pandas.DataFrame([[top_n, h_2_n, h_4_n, h_6_n, h_8_n, button_n, ave_n]], columns=columns)
            v = pandas.DataFrame([[velocity(top_e, top_n), velocity(h_2_e, h_2_n), velocity(h_4_e, h_4_n),
                                   velocity(h_6_e, h_6_n), velocity(h_8_e, h_8_n), velocity(button_e, button_n),
                                   velocity(ave_e, ave_n)]], columns=columns)
            d = pandas.DataFrame([[direction(top_e, top_n), direction(h_2_e, h_2_n), direction(h_4_e, h_4_n),
                                   direction(h_6_e, h_6_n), direction(h_8_e, h_8_n), direction(button_e, button_n),
                                   direction(ave_e, ave_n)]], columns=columns)
            self.ee = self.ee.append(e, ignore_index=True)
            self.nn = self.nn.append(n, ignore_index=True)
            self.vv = self.vv.append(v, ignore_index=True)
            self.dd = self.dd.append(d, ignore_index=True)
            all_depth.append(depth)
        self.depth = pandas.Series(all_depth)
        columns = ['time', 'top', 'h_2', 'h_4', 'h_6', 'h_8', 'button', 'ave']
        self.ee['time'] = self.df['t']
        self.ee = self.ee[columns]
        self.nn['time'] = self.df['t']
        self.nn = self.nn[columns]
        self.vv['time'] = self.df['t']
        self.vv = self.vv[columns]
        self.dd['time'] = self.df['t']
        self.dd = self.dd[columns]

    def save_result(self, outfile, V_D=True):
        if V_D:
            first = self.vv
            v_columns = {'top': 'v_0', 'h_2': 'v_2', 'h_4': 'v_4', 'h_6': 'v_6', 'h_8': 'v_8', 'button': 'v_10',
                         'ave': 'v'}
            second = self.dd
            d_columns = {'top': 'd_0', 'h_2': 'd_2', 'h_4': 'd_4', 'h_6': 'd_6', 'h_8': 'd_8', 'button': 'd_10',
                         'ave': 'd'}
            first = first.rename(columns=v_columns)
            second = second.rename(columns=d_columns)
            # emerge = pandas.concat([first, second],axis=1)
            columns = ['time', 'Depth', 'v_0', 'd_0', 'v_2', 'd_2', 'v_4', 'd_4', 'v_6', 'd_6', 'v_8', 'd_8', 'v_10',
                       'd_10', 'v', 'd']
        else:
            first = self.ee
            e_columns = {'top': 'e_0', 'h_2': 'e_2', 'h_4': 'e_4', 'h_6': 'e_6', 'h_8': 'e_8', 'button': 'e_10',
                         'ave': 'e'}
            second = self.nn
            n_columns = {'top': 'n_0', 'h_2': 'n_2', 'h_4': 'n_4', 'h_6': 'n_6', 'h_8': 'n_8', 'button': 'n_10',
                         'ave': 'n'}
            first = first.rename(columns=e_columns)
            second = second.rename(columns=n_columns)
            # emerge = pandas.concat([first,second],axis=1)

            columns = ['time', 'Depth', 'e_0', 'n_0', 'e_2', 'n_2', 'e_4', 'n_4', 'e_6', 'n_6', 'e_8', 'n_8', 'e_10',
                       'n_10', 'e', 'n']
        # emerge = emerge.ix[:, columns]
        # emerge = emerge.T.drop_duplicates().T
        emerge = pandas.merge(first, second, how='outer')
        emerge['Depth'] = self.depth
        emerge = emerge.reindex_axis(columns, axis=1)
        self.to_excel_style = emerge

        emerge.to_csv(outfile, sep=',', index=False, encoding='utf-8')

"""
a = r"E:\★★★★★项目★★★★★\★★★★★双子山油品储运贸易基地陆域形成工程★★★★★\实测数据\小潮\C6sun\C6VD.xlsx"
b = r"E:\★★★★★项目★★★★★\★★★★★双子山油品储运贸易基地陆域形成工程★★★★★\实测数据\小潮\C6sun\C6NE.csv"
c = Current(filename=r"E:\★★★★★CODE★★★★★\Git\debug\current_test\ne.csv",is_VD=False)
"""
# c.all_fenceng()
# c.save_result(r"E:\★★★★★项目★★★★★\★★★★★双子山油品储运贸易基地陆域形成工程★★★★★\实测数据\小潮\C6sun\C6result.csv")
"""

sheetname = 'ne'
c = current(filename=r"E:\★★★★★项目★★★★★\★★★★★双子山油品储运贸易基地陆域形成工程★★★★★\实测数据\大潮\C1刘鹏飞\ne.xlsx",sheet_name=sheetname,is_VD=False)
c.all_fenceng()
c.save_result(r"E:\★★★★★项目★★★★★\★★★★★双子山油品储运贸易基地陆域形成工程★★★★★\实测数据\大潮\C1刘鹏飞\处理结果.csv")
"""
"""
c.all_fenceng()
c.save_result(b)
"""


class One_Current_Point(object):
    def __init__(self, point, angle, ang=0, zhang_or_luo=False, cengshu=6):
        self.point = point
        self.angle = angle
        self.ang = ang
        self.zhang_or_luo = zhang_or_luo
        self.cengshu = cengshu

    def location(self, longitude=None, latitude=None, x=None, y=None):
        if longitude and latitude:
            self.longitude = longitude
            self.latitude = latitude
        if x and y:
            self.x = x
            self.y = y


class Single_Tide_Point(One_Current_Point):
    def __init__(self, filename, tide_type, point, angle, ang=0, zhang_or_luo=False, cengshu=6):  # False时，默认给出为落潮流向
        One_Current_Point.__init__(self, point, angle, ang, zhang_or_luo, cengshu)
        self.tide_type = tide_type
        self.filename = filename
        data = pandas.read_csv(filename)
        self.cengs = []
        if self.cengshu == 6:
            self.cengs.append(data[['time', 'v', 'd']])
            for i in range(0, 12, 2):
                self.cengs.append(data[['time', 'v_' + str(i), 'd_' + str(i)]])
                # 0为垂线，后面的依次是从上往下的层数
        self.ceng_processed = []
        for i in self.cengs:
            self.ceng_processed.append(time_v_d(i, self.angle, ang=self.ang, zhang_or_luo=self.zhang_or_luo))
        self.cal_time_of_ave()

    def cal_time_of_ave(self):
        self.time = self.cal_time(self.cengs[0])
        self.raise_time = self.time['last_time1']
        self.ebb_time = self.time['last_time2']
        self.first_ebb = False
        if self.zhang_or_luo:  # 给出为涨潮流向时
            if not self.cengs[0].loc[0]['timeof'] == 1:  # 初次为落潮
                self.first_ebb = True
                # else:
                #    First_ebb = False #初次为涨潮
        else:
            if self.cengs[0].loc[0]['timeof'] == 1:
                self.first_ebb = True
                # else:
                #    First_ebb  = False#初次为涨潮
        if self.first_ebb:
            self.raise_time, self.ebb_time = self.ebb_time, self.raise_time

    def cal_time(self, tvd):
        tvd.loc[tvd.index, 'timeof'] = tvd['d'].apply(lambda x: is_time_of(x)(self.angle, self.ang))
        tvd = add_convert_row_to_tvd_and_timeof(tvd, self.angle)
        return get_duration(tvd)

        # self.ave_of_vertical = self.cengs[0]
        # self.ave_of_vertical['timeof'] = self.ave_of_vertical['d'].apply(lambda x : is_time_of(x)(angle,ang))
        # self.ave_of_vertical = add_convert_row_to_tvd_and_timeof(***,ang)

        # self.time = get_duration(self.ave_of_vertical)

    def out_put(self, ):
        ceng_index = ['垂线平均', '表层', '0.2H', '0.4H', '0.6H', '0.8H', '底层', ]
        columns = ['层数', '平均流速', '平均流向', '最大流速', '最大流速对应方向', '出现时间']
        zhang = pandas.DataFrame(columns=columns)
        luo = pandas.DataFrame(columns=columns)
        for i, j in enumerate(ceng_index):
            ceng = self.ceng_processed[i]
            zhang = zhang.append(pandas.DataFrame(
                {"层数": j, "平均流速": ceng.zhang.mean, "平均流向": ceng.zhang.mean_d,
                 "最大流速": ceng.zhang.extreme_v, "最大流速对应方向": ceng.zhang.extreme_d, "出现时间": ceng.zhang.extreme_t}))
            luo = luo.append(pandas.DataFrame(
                {"层数": j, "最大流速": ceng.luo.extreme_v, "最大流速对应方向": ceng.luo.extreme_d,
                 "出现时间": ceng.luo.extreme_t, "平均流速": ceng.luo.mean,
                 "平均流向": ceng.luo.mean_d}))
        return zhang, luo

    def output_all(self):
        zhang, luo = self.out_put()
        zhang['涨落潮'] = '涨潮'
        luo['涨落潮'] = '落潮'
        statistics = zhang.append(luo, ignore_index=True)
        statistics['Point'] = self.point
        statistics['潮型'] = self.tide_type
        statistics['来源文件'] = self.filename
        return statistics

    def change_one_dir_values(self, timeof=1, parameter=0.8):
        changed = []
        for i in self.ceng_processed:
            changed.append(i.change_one_dir_values(timeof=timeof, parameter=parameter))
        for i in range(0, 6):
            changed[i + 1] = changed[i + 1].rename(columns={'v': 'v_' + str(i * 2), 'd': 'd_' + str(i * 2)})
        self.changed_out = pandas.concat(changed, axis=1)
        self.changed_out = self.changed_out.T.drop_duplicates().T
        columns = ['t', 'v_0', 'd_0', 'v_2', 'd_2', 'v_4', 'd_4', 'v_6', 'd_6', 'v_8', 'd_8', 'v_10', 'd_10', 'v', 'd']
        self.changed_out = self.changed_out[columns]
        self.changed_out = self.changed_out.reindex_axis(columns, axis=1)
        self.changed_out = self.changed_out.rename(columns={'t': 'time'})
        return self.changed_out

    def draw_dxf(self, parameter=10, ceng=0, drawing=None, *args):

        def plot_line(x, y, vs, ds, cengshu, parameter):
            for v, d in zip(vs, ds):
                line = dxf.line((x, y), end_point(x, y, v, d, parameter))

                drawing.add(line)
                # layer_name = dxf.layer(cengshu+self.tide_type)
                # drawing.layers.add(layer_name)
                line['layer'] = cengshu + self.tide_type

        def end_point(x, y, velocity, direction, parameter):
            v_east = velocity * np.sin(direction / 180 * np.pi)
            v_north = velocity * np.cos(direction / 180 * np.pi)
            return [x + v_east * parameter, y + v_north * parameter]

        data = self.ceng_processed[ceng].data
        data = data[data['t'].apply(lambda t: t.minute) == 0]
        if self.cengshu == 6:
            cengshu_name = ['垂线平均', '表层', '0.2层', '0.4层', '0.6层', '0.8层', '底层']
        else:
            cengshu_name = ['垂线平均', '表层', '中层', '底层']
        if not drawing:
            drawing = dxf.drawing(self.point + '流速矢量图.dxf')
        plot_line(self.x, self.y, data['v'], data['d'], cengshu_name[ceng], parameter)
        if not drawing:
            drawing.save()

    def display(self):
        e = []
        n = []
        timeof = []
        parameter = 0.95 / self.ceng_processed[0].data['v'].values.max()

        for i in self.ceng_processed:
            data = i.data
            e.append(east(data['v'] * parameter, data['d']).values)
            n.append(north(data['v'] * parameter, data['d']).values)
            timeof.append(data['timeof'])

        fig, ax = plt.subplots(1, 1, figsize=(len(data), 8))
        for i in range(1, 8):
            for j in range(1, len(data) + 1):
                if timeof[i - 1][j - 1] == -1:
                    color = 'green'
                else:
                    color = 'red'
                ax.arrow(j, i, e[7 - i][j - 1], n[7 - i][j - 1], head_width=0.1, head_length=0.1, fc=color, ec=color)
        ax.set_xlim(0, len(data))
        ax.set_ylim(0, len(self.ceng_processed) + 1)
        ax.set_xlabel('时间', fontsize=20)
        plt.yticks(range(8), ['', '底层', '0.8H', '0.6H', '0.4H', '0.2H', '表层', '垂线平均'], fontsize=20)
        plt.title(self.point + self.tide_type, fontsize=25)
        return fig

    def output_txt(self):
        out = self.output_all()


class time_v_d(object):
    def __init__(self, data, angle, ang=0, zhang_or_luo=False):  # angle为落潮流向,ang为转流区域角度
        for i in data.columns:
            if 'v' in i:
                # data['v'] = data[i]
                data.loc[data.index, 'v'] = data[i]
            if 'd' in i:
                data.loc[data.index, 'd'] = data[i]
            if 't' in i:
                # data['t'] = pandas.to_datetime(data[i])
                data.loc[data.index, 't'] = pandas.to_datetime(data[i])

        self.data = data[['v', 'd', 't']]
        self.data.loc[self.data.index, 'timeof'] = self.data['d'].apply(lambda x: is_time_of(x)(angle, ang))

        fenzhangluo = self.data.groupby('timeof')
        for ii, jj in fenzhangluo:
            if ii == 1:  #
                self.zhang = v_and_d_of_one_dir(jj)
            if ii == -1:
                self.luo = v_and_d_of_one_dir(jj)

    def change_one_dir_values(self, timeof=1, parameter=0.8):
        list_to_change = self.data[self.data['timeof'] == timeof]
        self.data.loc[list_to_change.index, 'v'] = list_to_change['v'] * parameter
        return self.data

class v_and_d_of_one_dir(object):
    def __init__(self, v_and_d):
        self.data = v_and_d
        self.extreme_v = v_and_d.v.max()
        self.extreme_d = v_and_d.loc[v_and_d['v'] == self.extreme_v]['d']
        t = (v_and_d.loc[v_and_d['v'] == self.extreme_v]['t'])
        self.extreme_t = t.values[0]
        # v_and_d['x'] = v_and_d['t'].apply(lambda x: x.minute)
        v_and_d.loc[v_and_d.index, 'x'] = v_and_d['t'].apply(lambda x: x.minute)
        v_and_d2 = v_and_d[v_and_d['x'] == 0]
        # v_and_d2['v_e'] = v_and_d2.apply(lambda df: east(df['v'], df['d']), axis=1)
        v_and_d2.loc[v_and_d2.index, 'v_e'] = v_and_d2.apply(lambda df: east(df['v'], df['d']), axis=1)
        # v_and_d2['v_n'] = v_and_d2.apply(lambda df: north(df['v'], df['d']), axis=1)10-23版本跟进pandas
        v_and_d2.loc[v_and_d2.index, 'v_n'] = v_and_d2.apply(lambda df: north(df['v'], df['d']), axis=1)
        self.mean = velocity(v_and_d2.v_e.mean(), v_and_d2.v_n.mean())  # 平均流速为矢量各分量平均之后的合成流速
        self.mean_d = direction(v_and_d2.v_e.mean(), v_and_d2.v_n.mean())


    def output_str(self, zhang_or_luo):
        if zhang_or_luo == "zhang":
            index = "涨潮"
        if zhang_or_luo == "luo":
            index = "落潮"
        result = ""
        result += index + '平均流速为' + str_from_float(self.mean), ',平均流向为' + str_from_float(self.mean_d)
        result += "\n"
        result += index + "流速最大为" + str_from_float(self.extreme_v), ',其对应流向为' + str_from_float(
            self.extreme_d.values[0]), '出现在' + str_from_datatime64(self.extreme_t)
        return result

def time_A2B(time_A, v_A, dir_A, time_B, v_B, dir_B, dir_cal):
    # dir_cal为转流经过的方向，在实际应用中应为涨/落潮方向垂直角度
    if not is_d1_in_d2_and_d3(dir_cal, dir_A, dir_B):
        dir_cal = dir_cal + 180 if dir_cal < 180 else dir_cal - 180

    ang_A = small_diff_dir(dir_A, dir_cal) / 180 * np.pi
    ang_B = small_diff_dir(dir_B, dir_cal) / 180 * np.pi
    projection_A = abs(v_A * np.sin(ang_A))
    projection_B = abs(v_B * np.sin(ang_B))
    delta_time = time_B - time_A
    mid_time = pandas.to_timedelta(projection_A / (projection_A + projection_B) * delta_time) + time_A
    return mid_time

def add_convert_row_to_tvd_and_timeof(time_v_d, angle):
    # angle为涨、落潮角度
    n = len(time_v_d)
    temp = pandas.DataFrame(columns=['t', 'v', 'd', 'timeof'])
    for i in range(0, n - 1):
        A = time_v_d.loc[i]
        B = time_v_d.loc[i + 1]
        if A.timeof * B.timeof == -1:
            t = time_A2B(A.t, A.v, A.d, B.t, B.v, B.d, dir_in_360(angle + 90))
            new_row = pandas.DataFrame({'t': [t - pandas.Timedelta(seconds=30), t + pandas.Timedelta(seconds=30)],
                                        'timeof': [A.timeof, B.timeof]})
            temp = temp.append(new_row, ignore_index=True)
    time_v_d = time_v_d.append(temp, ignore_index=True)
    time_v_d = time_v_d.sort_values(by='t')
    time_v_d['time'] = time_v_d['t']
    return time_v_d

c1  = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C1_d.csv", point = 'C1', tide_type='大潮', angle=275)
r"""

c1.location(x = 7.5596,y = 6.0510)

c1_x = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C1_x.csv",point = 'C1',tide_type='小潮',angle=275)

c2 = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C2_d.csv",point = 'C2',tide_type='大潮',angle=260)
c2_x = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C2_x.csv",point = 'C2',tide_type='小潮',angle=260)

c3 = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C3_d.csv",point = 'C3',tide_type='大潮',angle=280)
c3_x = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C3_x.csv",point = 'C3',tide_type='小潮',angle=280)

c4 = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C4_d.csv",point = 'C4',tide_type='大潮',angle=270)
c4_x = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C4_x.csv",point = 'C4',tide_type='小潮',angle=270)

c5 = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C5_d.csv",point = 'C5',tide_type='大潮',angle=270)
c5_x = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C5_x.csv",point = 'C5',tide_type='小潮',angle=270)

c6 = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C6_d.csv",point = 'C6',tide_type='大潮',angle=270)
c6_x = Single_Tide_Point(r"C:\Users\Feiger\Desktop\双子山潮流\C6_x.csv",point = 'C6',tide_type='小潮',angle=270)
print('*'*10)

x = c1.output_all()
a = [c1_x,c2,c2_x,c3,c3_x,c4,c4_x,c5,c5_x,c6,c6_x]
for i in a :
    x2 = i.output_all()
    x = x.append(x2,ignore_index=True)


#max_table = pandas.pivot_table(x,values = ['最大流速对应方向','最大流速'],index = ['Point','潮型','涨落潮'],columns = ['层数'])
#mean_table = pandas.pivot_table(x,values = ['平均流向','平均流速'],index = ['Point','潮型','涨落潮'],columns = ['层数'])
#tt = x.groupby(['Point','潮型','涨落潮'])
#x = x.set_index(['Point','潮型','涨落潮'])
"""
r"""for i,j in tt:
    print(i)
    v = j['最大流速'].max()
    print(v)
    t = j.loc[j['最大流速']==v]['出现时间']
    d = j.loc[j['最大流速']==v]['最大流速对应方向']
    ceng = j.loc[j['最大流速']==v]['层数']
    print('对应时间为'+str(t)+',最大流速对应方向为'+str(d)+',出现在层数'+str(ceng))
x.to_csv(r"C:\Users\Feiger\Desktop\分析结果.csv",sep=',')
#c0 = c.cengs[0]
#tvd = time_v_d(c0,275)"""


def draw_dxf(multi_point, parameter, filename):
    drawing = dxf.drawing('filename')
    for i in multi_point:
        i.draw_dxf(drawing=drawing)
    return drawing


print('OK')
r"""
c1_x.location(x = 7.5596,y = 6.5501)
c2_x.location(x = 4.2524,y = 5.6)
c2.location(x = 4.2524,y = 5.6)
c3.location(x = 2.6273,y = 4.801)
c3_x.location(x = 2.6273,y = 4.801)
c4.location(x = 5.3865,y = 5.106)
c4_x.location(x = 5.3865,y = 5.106)
c5.location(x = 4.3277,y = 3.791)
c5_x.location(x = 4.3277,y = 3.791)
c6.location(x = 7.5826,y = 3.8596)
c6_x.location(x = 7.5826,y = 3.8596)

d,x = [c1,c2,c3,c4,c5,c6],[c1_x,c2_x,c3_x,c4_x,c5_x,c6_x]
a = draw_dxf(d,parameter=0.05,filename='大潮流速矢量图.dxf')
a.save()
"""
