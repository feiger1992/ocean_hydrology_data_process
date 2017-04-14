#coding:utf-8
import pandas
import pickle
from dxfwrite import DXFEngine as dxf
import numpy as np
import datetime
def end_point(x,y,velocity,direction):
    v_east = velocity * np.sin(direction / 180 * np.pi)
    v_north = velocity * np.cos(direction / 180 * np.pi)
    return [x+v_east,y+v_north]

def get_data(filename ):
    print('数据来自文件'+filename)
    return pandas.read_excel(filename)

def plot_lines(data):
    def plot_line(x, y, vs, ds, cengshu):
        for v, d in zip(vs, ds):
            line = dxf.line((x, y), end_point(x, y, v, d))
            drawing.add(line)
            line['layer'] = cengshu
    style = data['tidestyle']
    drawing = dxf.drawing('流速矢量图.dxf')
    x = data['x'].dropna().values[0]
    y = data['y'].dropna().values[0]
    plot_line(x, y, data['v_0'].values, data['d_0'], '表层')
    plot_line(x, y, data['v_2'].values, data['d_2'], '0.2层')
    plot_line(x, y, data['v_4'].values, data['d_4'], '0.4层')
    plot_line(x, y, data['v_6'].values, data['d_6'], '0.6层')
    plot_line(x, y, data['v_8'].values, data['d_8'], '0.8层')
    plot_line(x, y, data['v_10'].values, data['d_10'], '底层')
    plot_line(x, y, data['v_max'].values, data['d_v_max'], '底层')
    plot_line(x, y, data['v_max'].values, data['d_v_max'], '底层')
    plot_line(x, y, data['v'].values, data['d'], '垂线平均')
    drawing.save()
    print('ok')

def get_time_of(angle1,angle2):
    def raise_or_ebb(d):
        if angle1<d<angle2:
            return 1
        else :
            return -1

def is_time_of(d):
    def raise_or_not(angle1):#涨/落潮主方向
        if angle1 > 180:
            raise ValueError
            return "Wrong with the angle"
        if angle1<90:
            angle1 = angle1 + 90
        else:
            angle1 = angle1 - 90
        angle2 = angle1 + 180
        if min(abs(d - angle1), abs(d - angle2), 360 - abs(d - angle2), 360 - abs(d - angle1)) < 30:
            return 0  # 转流
        elif angle1 < d < angle2:
            return 1  # 涨/落潮
        else:
            return -1 # 落/涨潮
    return raise_or_not

def get_extreme(v_and_d):#统计涨落潮极值
    extreme_v = v_and_d.v.max()
    extreme_d = v_and_d.loc[v_and_d['v'] == extreme_v]['d']
    return dict(v=extreme_v, d=extreme_d, mean=v_and_d.v.mean())

def get_duration(time_v_d):#统计涨落潮时间
    time_v_d = time_v_d[time_v_d.timeof != 0 ]
    time_v_d.index =  range(0,len(time_v_d))
    raise_time = 0  # 此时无法区分涨落，仅作标识用途
    ebb_time = 0
    raise_duration = datetime.timedelta(0)
    ebb_duration = datetime.timedelta(0)
    raise_last = False
    ebb_last = False
    raise_times = []
    ebb_times = []

    for i in range(1, len(time_v_d) ):
        if (time_v_d.loc[i, 'timeof'] == 1) and (time_v_d.loc[i - 1, 'timeof'] == 1) and raise_last:  # 涨潮last
            raise_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time'])
            continue

        if (time_v_d.loc[i, 'timeof'] == -1) and (time_v_d.loc[i - 1, 'timeof'] == -1) and ebb_last :  # 落潮last
            ebb_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time'])
            continue

        if (time_v_d.loc[i, 'timeof'] == -1) and (time_v_d.loc[i - 1, 'timeof'] == 1)  :  # 涨潮end,直接变落潮
            if raise_last:
                raise_time += 0.5
                raise_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time'])/2
                raise_last = False
                raise_times.append(raise_duration)
                raise_duration = datetime.timedelta(0)
            ebb_time += 0.5
            ebb_last = True
            ebb_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time']) / 2
            continue

        if (time_v_d.loc[i, 'timeof'] == 1) and (time_v_d.loc[i - 1, 'timeof'] == -1) :  # 落潮end,直接变涨潮
            if ebb_last:
                ebb_time += 0.5
                ebb_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time'])/2
                ebb_last = False
                ebb_times.append(ebb_duration)
                ebb_duration=datetime.timedelta(0)
            raise_time += 0.5
            raise_last = True
            raise_duration += (time_v_d.loc[i, 'time'] - time_v_d.loc[i - 1, 'time']) / 2
            continue
    print('raise_time=' + str(raise_time))#次数
    print('ebb_time=' + str(ebb_time))#次数
    return {'durations_1': raise_times,'durations_2': ebb_times,'times1':raise_time,'times2' :ebb_time,'last_time1':raise_duration/int(raise_time),'last_time2':ebb_duration/int(ebb_time) }

def statistics(data,angle1,zhang_or_luo = False):##zhang_or_luo为True时，angle为落潮流向
    fenceng = dict(biaoceng = data[[0,1,2]],
                  ceng2 = data[[0,3,4]],
                  ceng4=data[[0,5,6]],
                  ceng6=data[[0,7,8]],
                  ceng8=data[[0,9,10]],
                  diceng = data[[0,11,12]],
                   ave = data[[0,15,16]])
    max_1 = {}
    max_2 = {}
    for i in fenceng:
        fenceng[i].columns = ['time','v','d']
        fenceng[i]['timeof'] = (fenceng[i])['d'].apply(lambda x : is_time_of(x)(angle1))

        fenzhangluo = fenceng[i].groupby('timeof')
        print(str(i)+'__________________________________________')
        for ii,jj in fenzhangluo:
            if ii == 1: #
                max_1.update({i:get_extreme(jj)})#涨潮extreme

            if ii == -1:
                max_2.update({i:get_extreme(jj)})#落潮extreme
        print('极值1')
        print(max_1)
        print('极值2')
        print(max_2)
    if zhang_or_luo:
        time_of_raise = {'平均持续时间':get_duration(fenceng['ave'])['last_time1'],'涨落时间汇总':get_duration(fenceng['ave'])['durations_1']}
        time_of_ebb ={'平均持续时间':get_duration(fenceng['ave'])['last_time2'],'涨落时间汇总':get_duration(fenceng['ave'])['durations_2']}
        max_zhang = max_1
        max_luo = max_2
    else:
        time_of_ebb = {'平均持续时间':get_duration(fenceng['ave'])['last_time1'],'涨落时间汇总':get_duration(fenceng['ave'])['durations_1']}
        time_of_raise ={'平均持续时间':get_duration(fenceng['ave'])['last_time2'],'涨落时间汇总':get_duration(fenceng['ave'])['durations_2']}
        max_zhang = max_2
        max_luo = max_1
    return {'涨潮时间':time_of_raise,'落潮时间':time_of_ebb,'涨潮极值':max_zhang,'落潮极值':max_luo}
k = r"C:\Users\Feiger\Desktop\1.xls"
data = get_data(k)
haha = statistics(data,110)
#with open(r"C:\Users\Feiger\Desktop\输出\aaa.txt", 'wb') as file:
 #   pickle.dump(str(haha).encode('utf-8'), file)
print('汇总如下')
for i,j in haha.items():print(i,j)





