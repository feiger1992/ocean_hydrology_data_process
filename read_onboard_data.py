# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import os,sys
from collections import namedtuple
import matplotlib.pyplot as plt
import matplotlib
from scipy.interpolate import interp1d,splrep,splev
import warnings
import pycnnum
import lunardate
import win32com.client as client

sys.stdout = open(r"C:\Users\刘鹏飞\Desktop\test\logging.txt", 'w')

preamble = matplotlib.rcParams.setdefault('pgf.preamble', [])
preamble.append(r'\usepackage{color}')
velocity = lambda v_e, v_n: np.sqrt(v_e ** 2 + v_n ** 2)
dir_in_360b = lambda d: (d - 360 if d >= 360 else d) if d > 0 else 360 + d
dir_in_360 = lambda d: dir_in_360b(d) if (dir_in_360b(d) >= 0 and (dir_in_360b(d) < 360)) else dir_in_360b(
    dir_in_360b(d))
direction = lambda v_e, v_n: dir_in_360(180 / np.pi * np.arctan2(v_e, v_n)) if (v_e > 0) else dir_in_360(
    180 / np.pi * np.arctan2(v_e, v_n) + 360)
Coord = namedtuple('Coord','Latitude Longitude')
convertNum = lambda df: pd.to_numeric(df,errors = 'coerce').dropna().values
time_sub = lambda times,first_time: [(i -first_time).total_seconds()/3600 for i in times]
ss = lambda x: sorted(x, key=lambda o: o.time)

def gen_current_layer(times,Point_NAME = "",Project_NAME = "沈家湾客运中心三期码头工程水文测验",Project_AREA = "沈家湾",Line_NAME = "",Tide_Name = "潮汛：" ):
    Title = "附录：走航流速流向报表"
    Project_NAME = "工程名称："+Project_NAME
    Project_AREA = "海区：" + Project_AREA
    Line_NAME = "走航断面：" + Line_NAME
    Point_NAME = "测站：" + Point_NAME
    dates = describe_dates(times)
    arrays = [
                [Project_NAME] + [""]*13 + [dates],[""]*3 + [Point_NAME] + [""]*4 + [Tide_Name] + [""]*3 +["仪器型号："] + [""]*2 ,
        ["水深",'表层', '表层', '0.2H', '0.2H', '0.4H', '0.4H', '0.6H', '0.6H', '0.8H', '0.8H', '底层', '底层', '垂线平均', '垂线平均',
         ], ["m","(cm/s)", "(°)", "(cm/s)", "(°)", "(cm/s)", "(°)", "(cm/s)", "(°)", "(cm/s)", "(°)", "(cm/s)", "(°)", "(cm/s)", "(°)"]]
    tuples = list(zip(*arrays))
    current_layer = pd.MultiIndex.from_tuples(tuples, names=[Title, Project_AREA,'层数','时间'])
    return current_layer

def solar2lunar_d(d):
    # 将阳历转化为阴历
    # 汉字格式改写
    han_zi = pycnnum.num2cn(lunardate.LunarDate.fromSolarDate(d.year, d.month, d.day).day)
    han_zi = han_zi.replace('一十', '十')
    m =  pycnnum.num2cn(lunardate.LunarDate.fromSolarDate(d.year, d.month, d.day).month)
    m = m.replace("一十一","冬").replace("一十二","腊").replace("一十",'十') + "月"
    if len(m) == 1:
        if m == "一":
            m = "正"
    if len(han_zi) == 3:
        han_zi =  han_zi.replace('二十', '廿')
    if len(han_zi) == 1:
        han_zi = '初' + han_zi
    return m + han_zi

def describe_dates(times):
    a = min(times)
    b = max(times)
    solar = "观测时间:" +a.strftime("%y.%m.%d")+ "～" + b.strftime("%y.%m.%d")
    lunar = "农历" + solar2lunar_d(a) + "至" + solar2lunar_d(b)
    return solar + "(" + lunar + ")"

def aoe(fun, v_es, v_ns):
    vs = []
    for i, j in zip(v_es, v_ns):
        v = fun(i, j)
        vs.append(v)
    return vs

def distance(p1, p2):
    Lat = (p1.Latitude - p2.Latitude) * 111000
    Long = (p1.Longitude - p2.Longitude) * 111000 * np.cos(p1.Latitude)
    return int(np.sqrt(Lat ** 2 + Long ** 2))


def delete_blank_rows(excelfilename,n):
    #删除pandas保存bug引起的多余空行
    #其他格式调整

    exc = client.gencache.EnsureDispatch("Excel.Application")
    exc.Visible = 0
    wb = exc.Workbooks.Open(Filename=excelfilename)
    #wb.Worksheets(1).Columns("A").Select
    #exc.Selection.ColumnWidth = 25

    wb.Worksheets(1).Columns("A").ColumnWidth = 25
    for i in range(n, -1, -1):
        r = i * 40
        x = exc.Range("A%s:P%s" % (r + 1, r + 2))
        for border_id in range(7, 13):
            x.Borders(border_id).LineStyle = None
            # x.Borders(border_id).Weight=3
        x.Borders(9).LineStyle = 1

        exc.Range("B%s:P%s" % (r + 1, r + 2)).MergeCells = False

        wb.Worksheets(1).Cells(r + 1, 16).HorizontalAlignment = 4
        wb.Worksheets(1).Cells(r + 1, 2).HorizontalAlignment = 2

        exc.Rows("%d:%d" % (r+5, r+5)).Select()
        exc.Selection.Delete(Shift=client.constants.xlUp)
    exc.ActiveWorkbook.Save()
    exc.Quit()

class Single_time_point_data(object):
    def __init__(self,df):
        if df.iloc[5,1] != "cm" or df.iloc[5,2] != "BT":
            raise ValueError("数据格式有误")
        self.time = self.read_time(df)
        pos = convertNum(df.iloc[3,:2])
        self.position = Coord(pos[0],pos[1])
        self.depth = df.iloc[1,8:11].values.mean()
        self.v = convertNum(df.iloc[6:,1])
        self.d = convertNum(df.iloc[6:,2])
        self.e = convertNum(df.iloc[6:,3])
        self.n = convertNum(df.iloc[6:,4])
        self.larger_counts = len([i for i in self.v if i > 200])
        self.larger_counts_2 = len([i for i in self.v if i > 250 ])
        try:
            self.largest = max(self.v)
        except:
            pass
        #print( " at position of " + str(self.position))
    def read_time(self,df):
        row = df.iloc[0,:]
        row = [int(i) for i in row]
        row[0] += 2000
        row = [str(i) for i in row]
        s = row[0]+ '-'+row[1]+ '-'+row[2]+ ' '+row[3]+ ':'+row[4]+ ':'+row[5]+ '.'+row[6]
        #print("reading data of " +s ,end=",")
        return pd.Timestamp(s)

    def interpolate(self,x, v, x_new):
        f = interp1d(x, v)
        try:
            y = f(x_new)
        except:
            x_1 = x_new[0]
            tck = splrep(x,v,k = 1)
            y_1 = splev(x_1,tck).item()
            if abs(y_1) > 200:
                print("*"*20)
            x_2 = x_new[1:]
            try:
                y_2 = self.interpolate(x,v,x_2)
                y = [y_1] + list(y_2)
            except:
                y = splev(x_new,tck)
        return y

    def fenceng(self,datas,bin,blind_area):
        dep = bin * len(datas) + blind_area + bin / 2
        depths = [blind_area + bin]
        bottom = datas[-1] * 0.95
        if len(datas) > 2:
            top = datas[0]*1.05
            for i in range(1, len(datas)):
                depths.append(depths[-1] + bin)
            d2,d4,d6,d8 = dep * 0.2, dep * 0.4, dep * 0.6, dep * 0.8
            v2,v4,v6,v8 = self.interpolate(depths,datas,[d2,d4,d6,d8])
            return top, v2, v4, v6, v8, bottom
        if len(datas) == 2:
            v6 = datas[0] / 2 + datas[1] / 2
            return np.nan,datas[0]*1.05,np.nan,v6,bottom,np.nan

    def en_fenceng(self,bin,blind_area):
        self.e6 = self.fenceng(self.e, bin, blind_area)
        self.n6 = self.fenceng(self.n, bin, blind_area)
        self.e_init = self.cal_ave(self.e6)
        self.n_init = self.cal_ave(self.n6)
        self.init_v = velocity(self.e_init,self.n_init)
        self.init_d = direction(self.e_init, self.n_init)

    def cal_ave(self,datas):
        if pd.Series(datas).isnull().values.any():
            return (datas[1] +  2*datas[3] +datas[4])/4
        else:
            return (datas[0] + 2*datas[1] + 2*datas[2] + 2*datas[3] + 2*datas[4] + datas[5] )/10

    def set_pos_distance(self,distance):
        self.distance = distance

class Once_survey(object):
    def __init__(self,csv,position_file,bin,blind_area,IF_process = True,IF_Draw_raw_points = True, IF_Draw_selected_points = True):
        self.Df = pd.read_csv(csv,delim_whitespace = True,skiprows=3,header=None,na_values=-32768)
        os.chdir(r"C:\Users\刘鹏飞\Desktop\test")
        self.dfs = self.split_df(self.Df)
        self.time = csv[-23:-11]
        self.bin = bin
        self.blind_area = blind_area
        self.Collections = []
        self.invalid = []

        for i in self.dfs:
            try:
                ii = Single_time_point_data(i)
                if ii.position.Latitude == 30000:
                    self.invalid.append(ii)
                    print(str(ii.time) + "时无gps数据")
                elif ii.position.Longitude > 122.35:
                    self.invalid.append(ii)
                    print(str(ii.time) + "时gps数据偏差太大，为" + str(ii.position))
                else:
                    self.Collections.append(ii)
            except ValueError as Error_message:
                warnings.warn(csv + "文件中，第" + str(i) + "个读取的数据报错")
                warnings.warn(Error_message)
        self.Points = pd.read_csv(position_file)
        if IF_Draw_raw_points:
            self.draw_raw_points()
        if IF_process:
            self.select(IF_Draw_selected_points)

    def select(self,IF_Draw_selected_points):
        self.selected_points = {}
        self.selected_distance = {}
        self.selected_depth = {}
        self.check_points = {}

        def del_indices(indices):
            selected.v = np.delete(selected.v,indices)
            selected.d = np.delete(selected.d,indices)
            selected.e = np.delete(selected.e,indices)
            selected.n = np.delete(selected.n,indices)
            selected.larger_counts = len([i for i in selected.v if i > 200])
            selected.larger_counts_2 = len([i for i in selected.v if i > 250])
            selected.largest = max(selected.v)

        for p in self.Points.itertuples():
            self.check_points.update({p:False})
            selected = self.Collections[0]
            for run_points in self.Collections:
                if distance(p,run_points.position) < distance(p,selected.position) and len(run_points.v) > 1:
                    #距离近的点位可能是全部空值，因此加入后一项条件
                    self.check_points.update({p: True})
                    selected = run_points
                    dis = distance(p,selected.position)
                    self.selected_points.update({p.Name:selected})
                    self.selected_distance.update({p.Name:dis})
                    self.selected_depth.update({p.Name:round(selected.depth,1)})

            while ((selected.v[-1] > 100) and (selected.v[-1] > selected.v[-2] * 1.5) and (len(selected.v) > 3)) or \
                    ((selected.v[-1] > 120) and (selected.v[-1] > 2 * selected.v[0])  and (len(selected.v) < 10)) or \
                    ((selected.v[-1] > 50) and (selected.v[-1] > 2 * selected.v.mean())):
                print("点" + p.Name + "在" + selected.time.strftime("%m/%d %H:%M") + "时，底层流速偏大,为" + str(selected.v[-1])+ "m/s，将其删除." )
                del_indices(-1)
                print("删除之后其流速为"+ str(selected.v))

            while max(selected.e) > 220 :
                del_indices(np.argmax(selected.e))
                print("点" + p.Name + "在" + selected.time.strftime("%m/%d %H:%M") + "时，" + str(np.argmax(selected.e)) +"层流速偏大,为" + str(round(max(selected.e)))+ "m/s，将其删除." )
                print("删除之后其东向流速为" + str(selected.e))
            while max(selected.n) > 220:
                del_indices(np.argmax(selected.n))
                print("点" + p.Name + "在" + selected.time.strftime("%m/%d %H:%M") + "时，" + str(np.argmax(selected.n)) + "层流速偏大,为" + str(round(max(selected.n))) + "m/s，将其删除.")
                print("删除之后其北向流速为" + str(selected.n))
            while min(selected.e) < -220:
                del_indices(np.argmin(selected.e))
                print("点" + p.Name + "在" + selected.time.strftime("%m/%d %H:%M") + "时，" + str(np.argmin(selected.e)) + "层流速偏大,为" + str(round(min(selected.e))) + "m/s，将其删除.")
                print("删除之后其东向流速为" + str(selected.e))
            while min(selected.n) < -220:
                del_indices(np.argmin(selected.n))
                print("点" + p.Name + "在" + selected.time.strftime("%m/%d %H:%M") + "时，" + str(np.argmin(selected.n)) + "层流速偏大,为" + str(round(min(selected.n))) + "m/s，将其删除.")
                print("删除之后其北向流速为" + str(selected.n))

            if selected.largest > 250:
                print("点" + p.Name + "在"+ selected.time.strftime("%m/%d %H:%M") + "时，两米以上流速共出现" + str(selected.larger_counts) + "次，两米五以上流速共出现" + str(selected.larger_counts_2) + "次，" + "其流速数据为" + str(selected.v)+",其中最大流速为" + str(selected.largest) + "cm/s")

        for pp in self.selected_points:
            self.selected_points[pp].set_pos_distance(self.selected_distance[pp])
        for i in self.selected_points.values():
            i.en_fenceng(self.bin,self.blind_area)

        if IF_Draw_selected_points:
            self.draw_points()#默认处理过程全部出图

    def split_df(self,df):
        dfs = []
        n = int(df.iloc[5,0])
        for i in  range(int(len(df)/(n+6))):
            dfs.append(df.iloc[i*(n+6):(i+1)*(n+6),:])
        return dfs

    def draw_points(self):
        fig = plt.figure(figsize=(20, 20))
        ax = fig.add_subplot(1, 1, 1)
        ax.plot(self.Points.Longitude, self.Points.Latitude, 'b+')
        for p in self.Points.itertuples():
            s = p.Name + "^" + "{" + str(self.selected_distance[p.Name]) + "}" + "_{" + str(self.selected_depth[p.Name]) + "(" +str(len(self.selected_points[p.Name].v)) + ")}"
            s = "{" + s + "}"
            ax.annotate(r'$' + s + r'$', xy=(p.Longitude, p.Latitude))
        for name, selected_point in self.selected_points.items():
            ax.scatter(selected_point.position.Longitude, selected_point.position.Latitude, marker='*', c='r', s=10)
        plt.title(self.time)
        plt.tight_layout()
        plt.savefig(self.time + ".png",dpi = 400)
        plt.close()

    def draw_raw_points(self):
        self.raw_Longitudes = [i.position.Longitude for i in self.Collections]
        self.raw_Latitudes = [i.position.Latitude for i in self.Collections]
        fig = plt.figure(figsize=(20, 20))
        ax = fig.add_subplot(1, 1, 1)
        ax.plot(self.Points.Longitude, self.Points.Latitude,'r*' ,self.raw_Longitudes, self.raw_Latitudes, 'k+')
        ax.legend(["design_Points","survey_line"],loc = 'best')
        plt.title(self.time)
        plt.tight_layout()
        plt.savefig("All_Raw_Points"+self.time + ".png", dpi=400)
        plt.close()

class Process_One_Point_Data(object):
    def __init__(self,name,point_datas,If_Draw_interpolation = True):
        self.name = name
        point_datas = ss(point_datas)
        self.times, self.us, self.vs,self.positions,self.init_depths,self.distances,self.init_vs,self.init_ds = [], [], [],[], [], [],[],[]
        for once in point_datas:
            if once.distance < 80:
                self.distances.append(once.distance)
                self.times.append(once.time)
                self.us.append(once.e6)
                self.vs.append(once.n6)
                self.positions.append(once.position)
                self.init_depths.append(once.depth)
                self.init_vs.append(once.init_v)
                self.init_ds.append(once.init_d)
            else:
                warnings.warn("点" + name + "在" + str(once.time) + "时距离点位太远，距离为" + str(once.distance) + "米，点位已删除")
        self.us,self.vs = np.array(self.us),np.array(self.vs)
        print("Processing " + str(len(self.times)) + " datas of Point " + name )
        self.process(If_Draw_interpolation,self.times,self.us,self.vs,self.positions,self.init_depths,self.distances,name)

    def process(self,If_Draw_interpolation, times, us, vs, positions, depths,distances,name):
        #self.Us = Process_Velocity_Data(times,us,method=2)
        #self.Vs = Process_Velocity_Data(times,vs, method=2)
        print("东分量：" + "-" * 10)
        self.Us = Process_Velocity_Data(times, us, method=2)
        print("北分量：" + "-" * 10)
        self.Vs = Process_Velocity_Data(times, vs, method=2)
        if If_Draw_interpolation:
            self.Us.draw_interpolate_2('Velocity_Interpolation_' + name ,self.Vs)

        self.new_us,self.new_vs = self.Us.datas,self.Vs.datas
        x = self.Us.x
        x_new = self.Us.x_new
        self.cal_times = self.Us.cal_times

        tck2 = splrep(x,depths,k = 1)
        self.depths = splev(x_new,tck2)

        self.positions = pd.DataFrame(times, positions)
        self.distances = pd.DataFrame(times,distances)
        self.v = aoe(velocity, self.Us.ave, self.Vs.ave)
        self.d = aoe(direction, self.Us.ave, self.Vs.ave)

    def out_put(self,excel_file,Excel_writer = None,sheet_name = None, start_row = None):

        self.outDatas = []
        self.outDatas.append(self.depths.tolist())
        for i in range(6):
            self.outDatas.append(aoe(velocity,self.Us.datas[i],self.Vs.datas[i]))
            self.outDatas.append(aoe(direction, self.Us.datas[i], self.Vs.datas[i]))
        self.outDatas.append(self.v)
        self.outDatas.append(self.d)
        df = pd.DataFrame(self.outDatas).T

        self.outDF = pd.DataFrame(df.values,index = self.cal_times,columns=gen_current_layer(self.cal_times,self.name))
        if not Excel_writer:
            Excel = pd.ExcelWriter(excel_file,index_label = None)
            self.outDF.to_excel(Excel)
        else:
            #self.outDF.to_excel(excel_writer= Excel_writer,sheet_name= sheet_name,index_label="Time",float_format='%.3f',startrow=start_row)
            self.outDF.to_excel(excel_writer=Excel_writer, sheet_name=sheet_name,
                                float_format='%.3f', startrow=start_row,index_label=None)
    def display(self):
        self.display_datas(self.cal_times,self.v,self.d,self.depths,title = self.name)

    def display_init(self):
        self.display_datas(self.times,self.init_vs,self.init_ds,self.init_depths,title = "Init_" + self.name )

    def display_datas(self,time,v,d,depths,title ):
        matplotlib.rcParams.update({'font.size': 20})
        def make_patch_spines_invisible(ax):
            ax.set_frame_on(True)
            ax.patch.set_visible(False)
            for sp in ax.spines.values():
                sp.set_visible(False)
        fig, host = plt.subplots(figsize=(20, 10),dpi = 100 )
        fig.subplots_adjust(right=0.75)
        par1 = host.twinx()
        par2 = host.twinx()
        par2.spines["right"].set_position(("axes", 1.1))
        make_patch_spines_invisible(par2)
        par2.spines["right"].set_visible(True)

        p1, = host.plot(time, v, "b--", label="Velocity")
        p2, = par1.plot(time, d, "r-.", label="Direction")
        p3, = par2.plot(time, depths, "k-", label="Depths")


        par1.set_ylim(0, 360)

        host.set_xlabel("Time")
        host.set_ylabel("Velocity(cm/s)")
        par1.set_ylabel("Direction(°)")
        par2.set_ylabel("Depth(m)")

        host.yaxis.label.set_color(p1.get_color())
        par1.yaxis.label.set_color(p2.get_color())
        par2.yaxis.label.set_color(p3.get_color())

        tkw = dict(size=4, width=1.5)
        host.tick_params(axis='y', colors=p1.get_color(), **tkw)
        par1.tick_params(axis='y', colors=p2.get_color(), **tkw)
        par2.tick_params(axis='y', colors=p3.get_color(), **tkw)
        host.tick_params(axis='x', **tkw)
        lines = [p1, p2, p3]
        host.legend(lines, [l.get_label() for l in lines])

        hfmt = matplotlib.dates.DateFormatter('%H:%M')
        host.xaxis.set_major_formatter(hfmt)

        dates_span = min(time).strftime("%b/%d") +"-" +max(time).strftime("%b/%d")
        plt.title("Varying of Velocity,Direction and Depth at Point " + self.name + "(" + dates_span + ")")
        #plt.subplots_adjust(left=0.5, bottom=0.5, top=0.9)
        plt.savefig( title + " V_D_depth.png")
        plt.close()

class Process_Velocity_Data(object):
    def __init__(self, times, array_values, method=2):
        # array_values 用第二个下标循环层数
        self.get_times = times
        self.init_data = array_values.T
        self.cal_times = self.cal_time(self.get_times)
        self.datas, self.datas_all = [], []
        # datas all is set for plot figure
        self.x = time_sub(self.get_times, min(self.cal_times))
        self.x_new = time_sub(self.cal_times, min(self.cal_times))
        self.x_all = np.linspace(self.x[0],self.x[-1],500)
        self.three_storey = False
        for values in self.init_data:
            series = pd.Series(data=values)
            if not series.isnull().values.any():
                self.y = self.interpolate_2(self.x, values, self.x_new)
                if max(values) > 200:
                    print("---" * 10)
                    print(pd.Series(np.around(values.astype(np.double),3),index=self.get_times))
                    print("***" * 10)
                    print(pd.Series(np.around(self.y.astype(np.double),3), index=self.cal_times))
                    print("---"*10)
                self.datas.append(self.y)
                self.datas_all.append(self.interpolate_2(self.x, values, self.x_all))
            else:
                self.three_storey = True
                y = np.full(len(self.x_new),np.nan)
                y2 = np.full(len(self.x_all),np.nan)
                self.datas.append(y)
                self.datas_all.append(y2)
        if not self.three_storey:
            self.ave = self.datas[0] + 2* self.datas[1] +2* self.datas[2]+2* self.datas[3]+2* self.datas[4]+self.datas[5]
            self.ave = [i/10 for i in self.ave]
            self.datas = np.array(self.datas)
        else:
            self.ave = self.datas[1] + 2*self.datas[3] + self.datas[4]
            self.ave = [i/4 for i in self.ave]

    def draw_interpolate_2(self, title,datas2):
        #ax = fig.add_subplot(1, 1, 1)
        matplotlib.rcParams.update({'font.size': 12})
        fig, axs = plt.subplots(6,2, sharex=True, sharey=True, gridspec_kw={'hspace': 0})
        fig.suptitle(title)
        def hline(i):
            for index in [0,1]:
                axs[i, index].axhline(y=200, linewidth=0.1, color='blue')
                axs[i, index].axhline(y=-200, linewidth=0.1, color='blue')
                axs[i, index].axhline(y=250, linewidth=0.1, color='red')
                axs[i, index].axhline(y=-250, linewidth=0.1, color='red')
        for i in range(6):
            axs[i,0].plot(self.x, self.init_data[i], 'go',markersize=1)
            axs[i,0].plot(self.x_new, self.datas[i], 'r*', markersize=1)
            axs[i,0].plot(self.x_all, self.datas_all[i], 'y--',linewidth=0.5)
            axs[i,0].label_outer()
            axs[i,1].plot(self.x, datas2.init_data[i], 'go',markersize=1)
            axs[i,1].plot(self.x_new, datas2.datas[i], 'r*', markersize=1)
            axs[i,1].plot(self.x_all, datas2.datas_all[i], 'y--',linewidth=0.5)
            axs[i,1].label_outer()
            hline(i)

        axs[0,0].set_title("East(u)")
        axs[0,1].set_title("North(v)")
        fig.legend(['datas', "cal_ones", "ALL"],ncol=3, loc='lower center')
        plt.savefig(title + ".png", dpi=400)
        plt.close()

    def cal_time(self, times):
        t1, t2 = min(times), max(times)
        t1a, t2 = t1.ceil('h'), t2.floor('h')
        cal_times = pd.DatetimeIndex(start=t1a, end=t2.floor('h'), freq='h')
        if not len(cal_times) == 28:
            cal_times = pd.DatetimeIndex(start=t1.round('H'), end=t2.floor('h'), freq='h')
        return cal_times

    def interpolate_2(self, x, v, x_new):
        tck = splrep(x,v,k = 2)
        y = splev(x_new,tck)

        tck2 = splrep(x,v,k = 1)
        y2 = splev(x_new,tck2)
        c = 0
        try:
            while((max(y) > max(v) * 1.3) or (max(y) > 200)):
                y[np.argmax(y)] = y2[np.argmax(y)] / 2 + max(y)/2
                c +=1
                if c > 20:
                    break
            c = 0
            while((min(y) < min(v) * 1.3)  or (min(y)< -200)) :
                y[np.argmin(y)] = y2[np.argmin(y)] / 2 + min(y)/2
                c +=1
                if c > 20:
                    break
            return y
        except:
            pass

class Tide_survey(object):
    def __init__(self,csv_files,pos_file,outExcel,IF_process = False,IF_Draw_raw_points = False, IF_Draw_selected_points = False,If_Draw_interpolation = False,if_Draw_Arrows_Others = False,layer_height = 1,blank_area = 2.3):
        datas = []
        for i in  csv_files:
            print(" - - " * 10 +"\n"+"reading datas from " + i)
            Once = Once_survey(i,pos_file,layer_height,blank_area,IF_process ,IF_Draw_raw_points, IF_Draw_selected_points)#默认为洋山走航设置，盲区为2.3m，层高1m
            print("- - " * 10)
            datas.append(Once.selected_points)
        self.DF = pd.DataFrame(datas)
        self.ALL_datas = {}
        excelWriter = pd.ExcelWriter(outExcel)
        row = 0
        for p in self.DF.columns:
            point_datas = self.DF[p]
            P = Process_One_Point_Data(p,point_datas,If_Draw_interpolation)
            P.out_put(excelWriter,excelWriter,sheet_name="走航流速报表",start_row=row)
            row = row + 40
            P.display()
            P.display_init()
            self.ALL_datas.update({p:P})
            print("- * OK * -" * 10)

        excelWriter.close()
        delete_blank_rows(outExcel,len(self.DF.T))

        print("OK" * 10)
        if if_Draw_Arrows_Others:
            self.Points = pd.read_csv(pos_file)

            uDF,vDF = [],[]
            for p in self.Points.itertuples():
                p_data = self.ALL_datas[p.Name]
                uDF.append(np.round(p_data.Us.ave,1))
                vDF.append(np.round(p_data.Vs.ave,1))
            uu = pd.DataFrame(uDF, index=self.Points.Name, columns= P.cal_times)
            vv = pd.DataFrame(vDF, index=self.Points.Name, columns= P.cal_times)

            for time in P.cal_times:
                title = time.strftime("%m-%d %H%M")
                fig = plt.figure(figsize=(20, 20))
                ax = fig.add_subplot(111)
                for i in self.Points.itertuples():
                    u= uu.loc[i.Name, time]
                    v= vv.loc[i.Name,time]
                    V = round(velocity(u,v),1)
                    D = round(direction(u,v),1)
                    ax.text(i.X, i.Y, i.Name, color='green')
                    #ax.text(i.X + u , i.Y + v, "v = " + str(V) + "\nd = " + str(D), color='red', size="x-small")
                    ax.arrow(i.X, i.Y,u ,v , width=3, fc="black", ec="ivory")
                ax.set_xlim(512000, 513342)#
                ax.set_ylim(3386066, 3386910)#
                plt.title(str(time))
                plt.tight_layout()
                plt.savefig(title + ".png")
                plt.close()

            for i in self.Points.itertuples():
                fig = plt.figure(figsize=(20, 20))
                ax = fig.add_subplot(111)
                for time in P.cal_times:

                    u = uu.loc[i.Name, time]
                    v = vv.loc[i.Name, time]
                    ax.arrow(0, 0, u, v, width=3, fc="black", ec="ivory")
                    ax.set_xlim(-200, 200)
                    ax.set_ylim(-200, 200)
                plt.title(i.Name)
                plt.tight_layout()
                plt.savefig("ALL_Current_Data_of" + i.Name + ".png")
                plt.close()

            fig = plt.figure(figsize=(20, 20))
            ax = fig.add_subplot(111)
            for time in P.cal_times:
                for i in self.Points.itertuples():
                    u= uu.loc[i.Name, time]
                    v= vv.loc[i.Name,time]
                    V = round(velocity(u,v),1)
                    D = round(direction(u,v),1)
                    ax.arrow(i.X, i.Y,u ,v , width=1, fc="black", ec="ivory")
            ax.set_xlim(512000, 513342)#x界限
            ax.set_ylim(3386066, 3386910)#y界限
            plt.title("ALL_Current_Data")
            plt.tight_layout()
            plt.savefig("ALL_current_data" + ".png")
            plt.close()

da =[r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704091125_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704102435_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704110015_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704122247_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704130056_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704142423_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704150017_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704162722_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704170013_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704182952_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704190042_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704202651_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704210012_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704222818_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190704225836_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705002325_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705005629_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705022717_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705025819_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705042750_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705050304_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705062754_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705070006_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705082727_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705090000_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705102839_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705110011_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705123209_000t.000",r"C:\2019-沈家湾三期（GK-2019-0118水）\实测数据\current-onboard\大潮\大潮结果\xysdc20190705130524_000t.000"]
pos = r"C:\2019-沈家湾三期（GK-2019-0118水）\走航点位坐标.csv"
#Xiao = Tide_survey(xiao,pos,r"小潮合并.xlsx",IF_process=True,IF_Draw_raw_points=False,IF_Draw_selected_points=False,If_Draw_interpolation=False)
Da = Tide_survey(da,pos,r"C:\Users\刘鹏飞\Desktop\test\大潮合并aaa.xlsx",IF_process=True,IF_Draw_raw_points=False,IF_Draw_selected_points=False,If_Draw_interpolation=False,if_Draw_Arrows_Others = False)
#Zhong = Tide_survey(zhong,pos,r"中潮合并.xlsx",IF_process=True,IF_Draw_raw_points=False,IF_Draw_selected_points=False,If_Draw_interpolation=False)

print('OK! ' * 10)