from windrose import WindroseAxes
from matplotlib import pyplot  as plt
import matplotlib.cm as cm
import logging
import pandas as pd
import numpy as np
from os import chdir
from numpy import inf,histogram2d
import win32com.client as client

logging.basicConfig(level=logging.INFO,format='%(levelname)s %(asctime)s %(message)s',datefmt='%m/%d/%Y %H:%M:%S')
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
dir_in_360b = lambda d: (d - 360 if d >= 360 else d) if d > 0 else  360 + d
dir_in_360 = lambda d: dir_in_360b(d) if (dir_in_360b(d) >= 0 and (dir_in_360b(d) < 360)) else dir_in_360b(
    dir_in_360b(d))
EWSN2d = { "NNW":337.5, "NW":315, "WNW":292.5, "W":270, "WSW":247.5, "SW":225, "SSW":202.5, "S":180, "SSE":157.5, "SE":135, "ESE":112.5,"E":90,"ENE":67.5,"NE":45,"NNE":22.5,"N":0}
seasons = ['冬天', '冬天', '春天', '春天','春天', '夏天', '夏天','夏天', '秋天','秋天', '秋天', '冬天']
DIRECTIONS = ['E', 'ENE', 'NE', 'NNE', 'N', 'NNW', 'NW', 'WNW', 'W', 'WSW', 'SW', 'SSW', 'S', 'SSE', 'SE','ESE']
DIRECTIONS_N1 = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW","NNW"]
month_to_season = dict(zip(range(1,13), seasons))
is_time_start = lambda time: time.hour == 0 and time.is_month_start
is_time_end = lambda time: time.hour == 23 and time.is_month_end
theta_angles = np.arange(0, 360, 45/2)/180*np.pi
def close_line(ax):
    for line in ax.lines:
        x, y = line.get_data()
        if x[0] != x[-1]:
            x = np.concatenate((x, [x[0]]))
            y = np.concatenate((y, [y[0]]))
            line.set_data(x, y)

def insert_x_to_list(x,l): #l必须为单调递增
    #从L最左往右的区间挨个数，第一个区间为（-∞，l[0]]#0，后续每个区间为：(l[i],l[i+1])左开右闭#（i+1） ,最后一个区间为全开(l[-1],+∞)
    for i in range(1,len(l)):
        if l[i - 1] < x <= l[i]:
            return (i)
        else:
            if x <= l[0]:
                return 0
            elif x > l[-1]:
                return len(l)

def one_year_df(df):
    WHOLE_ANNUAL_DAYS = 365
    if 2 in df.index[np.where(df.index.is_leap_year == True)].month:
        WHOLE_ANNUAL_DAYS = 366
    if len(df) == np.where(df.index.is_month_end == True)[0][-1] + 1:
        df = df.iloc[np.where(df.index.is_month_start == True)[0][0]:, :]
    else:
        df = df.iloc[np.where(df.index.is_month_start == True)[0][0]:(
                np.where(df.index.is_month_end == True)[0][-1] + 1), :]
    d = df.index[-1] - df.index[0]
    if d.days == WHOLE_ANNUAL_DAYS - 1  and d.seconds/3600 == 23:
        logging.info('确认 %s 为全年统计开始时间\n %s 为全年统计结束时间,总历时为%s 天另%s小时' % (df.index[0], df.index[-1],WHOLE_ANNUAL_DAYS,d.seconds/3600))
    return df

def histogram222(x, y, bins):
    bins_x, bins_y = bins
    #print(len(bins_x))
    #print(len(bins_y))
    nx, ny = len(bins_x) , len(bins_y)
    hist = np.zeros((nx, ny))

    for i, j in zip(x, y):
        #print(i,"插入序数为",insert_x_to_list(i,bins_x))
        #print(j,"插入序数为",insert_x_to_list(j,bins_y))
        hist[insert_x_to_list(i, bins_x), insert_x_to_list(j, bins_y)] += 1
    #print(hist.shape)
    return hist

def d2EWSN(d):
    if d > 360 or d < 0:
        return d2EWSN(dir_in_360(d))
    if 348.75<d<360:
        return 'N'
    elif d >= 326.25:
        return "NNW"
    elif d >= 303.75:
        return "NW"
    elif d > 281.25:
        return "WNW"
    elif d >= 258.75:
        return "W"
    elif d > 236.25:
        return "WSW"
    elif d >= 213.75:
        return "SW"
    elif d >= 191.25:
        return "SSW"
    elif d >= 168.75:
        return "S"
    elif d >= 146.25:
        return "SSE"
    elif d >= 123.75:
        return "SE"
    elif d >= 101.25:
        return "ESE"
    elif d >= 78.75:
        return "E"
    elif d >= 56.25:
        return "ENE"
    elif d >= 33.75:
        return "NE"
    elif d >= 11.25:
        return "NNE"
    elif d <11.25:
        return "N"

def plot_data_of_one_item(var, dir_a, item, ang=40, rmax=None,bins=None, N=None, unit=None, title=None,
                          fontsize=None,per = None):
    # var 为所描述数值
    # bins 为单个条目的概率区分数列，np.linspace(...)，传入整数时默认为分级个数，默认为5
    # N为方向的区分度，有几圈,默认为5
    # ang 为标注各圈概率值的标签所在数轴角度
    # unit 为所分布值的单位
    # rmax 为最大概率值，即最外圈所表示的百分比
    bins = bins[:-1]

    fig = plt.figure(figsize=(16, 9), dpi=200)
    ax = WindroseAxes.from_ax(fig=fig, rmax=rmax)
    ax.bar(dir_a, var, normed=True, opening=0.8, edgecolor='white', bins=bins, N=N, cmap=cm.rainbow)
    # ax.box(dir, var, normed=True, edgecolor='white', bins=bins,  cmap=cm.rainbow)
    ax.set_legend(title=item + unit, fancybox=False, facecolor='ivory', edgecolor='black',
                  ncol=bins // 5 if isinstance(bins, int) else len(bins) // 5
                  , fontsize=15, bbox_to_anchor=(0, 0))
    ax.set_radii_angle(angle=45 if not ang else ang)
    ax.tick_params(axis='y', direction='inout', colors='darkblue', pad=1)
    ax.tick_params(axis='x', colors='black', labelsize=15, pad=2)
    ax.grid(color='black', linestyle=':', linewidth=1, alpha=1)
    ax.set_rorigin(-rmax/11)
    ax.text(.5,.5,"缺测\n数据\n" + str( round(100 -float(per),2))+"%",horizontalalignment='center',verticalalignment='center', transform=ax.transAxes)

    title = item + " " + title if title else item
    ax.set_title(title + "(有效数据："+ per[:5] + "%)", {'fontsize': 20 if not fontsize else fontsize})
    fig_file = title + '.png'
    fig.savefig(fig_file.replace(r'/', '_'), dpi=200, bbox_inches='tight')
    logging.info(fig_file.replace(r'/', '_') + '保存成功')
    plt.close()

def plot_distribution_of_df(group,key,title,N = 5):#N是圈数
    df = group.describe()[key][['max', 'min', 'mean',  '75%','50%','25%','count']]
    df = df.reindex(DIRECTIONS,fill_value=0)
    fig = plt.figure(figsize=(9, 9), dpi=200)
    ax = fig.add_subplot(111, projection='polar')
    ax.tick_params(axis='x',which = 'major', labelsize = 15)

    colors = ['k', 'w', '#7ef9ff']
    #c2 = ['#73c2fb', '#0080ff', '#010080']
    c2 = ['#ccccff', '#8080ff', '#4d4dff']
    for i, c1 in enumerate(c2):
        i += 3
        ax.fill(theta_angles, df.iloc[:, i].values, color=c1, label=df.columns[i])
    for i, c in enumerate(colors):
        ax.plot(theta_angles, df.iloc[:, i].values, color=c, label=df.columns[i], lw=1.5)

    close_line(ax)
    #ax.plot([],[],' ',label = "DIR(count)")

    theta_show_labels = []
    for j in DIRECTIONS:
        theta_show_labels.append(j + "(" + str(int(df.loc[j,'count'])) + ")")
    #y1 = np.floor(df['min'].min())
    y1 = 0
    y2 = np.ceil(df['max'].max())
    ax.set_xticks(theta_angles)
    ax.set_xticklabels(theta_show_labels)
    ax.set_rlabel_position(45 / np.pi)
    ax.set_rorigin(y1)

    ax.set_ylim(y1,y2)
    radii = np.linspace(y1,y2,num=6)
    ax.set_yticks(radii[1:])

    if y2 % N == 0:
        ax.yaxis.set_major_formatter(plt.FuncFormatter('{:.0f}'.format))
    else:
        ax.yaxis.set_major_formatter(plt.FuncFormatter('{:.1f}'.format))
    ax.grid(True,'major','y',lw = .5)

    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05),
              fancybox=True,  ncol=7,facecolor='#d4dade')
    ax.yaxis.label.set_color("#107c41")

    title = title + "（N = " + str(int(df['count'].sum())) + "）"
    ax.set_title(title, {'fontsize': 17})
    fig_file = title + '.png'
    fig.savefig(fig_file.replace(r'/', '_'), dpi=200, bbox_inches='tight')
    logging.info(fig_file.replace(r'/', '_') + '保存成功')
    plt.close()

class Data(object):
    def __init__(self,filename,draw_type,isEWSN = True, if_draw = True,if_statics = True,workdir = None,*args,**kwargs):

        try:
            self.sitename = kwargs['sitename']
            kwargs.pop('sitename',None)
            logging.info('开始读取文件 %s',filename)
            df = pd.read_excel(filename, index_col='time', converters={'time': pd.to_datetime})
            df = df.rename({'DIR':'dir','Dir':'dir','D':'dir','d':'dir'},axis = 1)
            df = df.set_index(df.index.round('min'))

            logging.info("读取文件成功，共有%s列，列标题为：%s",len(df.columns),df.columns.values)
            logging.info("数据时间从%s开始至%s结束",df.index[0],df.index[-1])
            self.rows = len(df)
            s = list(df.columns)
            s.remove('dir')
            for c_name in s:
                df[c_name] = pd.to_numeric(df[c_name], errors="coerce")

            df = df.dropna()
            logging.info("原始数据有%s行内容，其中有效数据为%s行",self.rows, len(df))
        except ValueError as e:
            logging.error('时间列发生错误:%s',e)
            raise Exception("时间错误")
        if ('dir' or 'DIR' or "Dir") not in df.columns:
            logging.error('***** 请将方向列标题改为“dir” *****')
            raise Exception("列标题错误")
        if isEWSN:
            df['DIR'] = df['dir']
            df['dir'] = df['dir'].apply(lambda a: EWSN2d.get(a))#dir是数值，DIR是字母
        else:
            df['DIR'] = df['dir'].apply(d2EWSN)
        df = df.rename({'dir': 'd'}, axis=1)

        df['month'] = df.index
        df['month'] = df['month'].apply(lambda x: (x.year, x.month))
        self.g_m = df.groupby('month')
        logging.info("共有%s个月的数据",len(self.g_m))

        df_year = one_year_df(df)
        df_year['season'] = df_year.index.month
        df_year['season'] = df_year['season'].apply(lambda x: month_to_season.get(x))
        self.g_s = df_year.groupby('season')

        self.df = df_year
        self.max_p = {}

        if if_draw:
            self.draw(draw_type,*args,**kwargs)

    def draw(self,draw_type,key,bins,*args,**kwargs):

        self.statics(key,bins)
        rmax = np.ceil(max(self.max_p[key])*20)*5
        unit = '(' + kwargs['unit'] + ')'
        kwargs['unit'] = unit

        if 1 in draw_type :
            self.draw_annual_data(rmax = rmax,bins =bins,key = key,*args,**kwargs)

        if 2 in draw_type:
            self.draw_month_data(rmax = rmax,bins =bins,key = key,*args,**kwargs)

        if 3 in draw_type:
            self.draw_season_data(rmax = rmax,bins =bins,key = key,*args,**kwargs)
    
    def extract_data(self,df,key):
        try:
            data,dir_data = df[key],df['d']
            logging.info("共找到有效数据%s个",len(data))
            return data,dir_data
        except KeyError as e:
            logging.error("文件数据中未找到对应列名称,%s",e)

    def count_df(self,df):
        start,end = df.index[0],df.index[-1]
        if not is_time_start(start):
            start = start.floor('D')
            start = start + pd.DateOffset(1 - start.day)
        if not is_time_end(end):
            end = end.ceil('D')
            end = end + pd.DateOffset(end.daysinmonth - end.day)

        count = (end - start ).total_seconds()
        count = round(count /(df.index[1] - df.index[0]).total_seconds() ) + 1
        logging.info("总过程应有%s个数据,现有%s个" %(count,len(df)))
        return count

    def distribution(self,df,key,v_bin,count = None):
        column_name = self.column_name(v_bin)
        d_bin = [11.25,  33.75,  56.25,  78.75, 101.25, 123.75, 146.25,
       168.75, 191.25, 213.75, 236.25, 258.75, 281.25, 303.75, 326.25,
       348.75, 371.25]
        v, d = self.extract_data(df,key)
        Max = df.groupby('DIR')[key].max()
        Mean = df.groupby('DIR')[key].mean()

        #h, e1, e2 = histogram2d(v, d, [v_bin, d_bin])#左闭右开，numpy自带函数
        h = histogram222(v,d,[v_bin, d_bin])#左开右闭，自己修改,慎调
        e1, e2 = v_bin,d_bin
        h = h[1:]#移除第一行，调整格式；为防止漏掉v中等于v_bin第一项的情况，输入v_bin中第一项应为0
        h[:,0] = h[:,0] + h[:,-1]
        h = h[:,:-1]#合并N的两行为一行
        count = self.count_df(df)

        df = pd.DataFrame(data=h/count, index=e1[:-1], columns=e2[:-1])
        df.columns = DIRECTIONS_N1
        df = df.T.reindex(DIRECTIONS).T
        df['缺测'] = 0
        df['汇总'] = df.T.sum()
        df = df.T
        df.columns = column_name
        df['汇总'] = df.T.sum()
        m = max(df['汇总'][:-1])
        self.max_p[key].append(m)

        df.loc['缺测','汇总'] = 1 - df.loc['汇总','汇总']
        df.loc['汇总', '汇总'] = 1

        df *= 100
        df['最大值'] = Max
        df['平均值'] = Mean
        df.loc['汇总','最大值'] = df['最大值'].max()
        df.loc['汇总', '平均值'] = df['平均值'].mean()
        df = df.fillna(0)
        df[df == 0] = "-"
        return df

    def column_name(self,bins):
        names = []
        for i in range(len(bins) - 1):
            names.append("(" + str(bins[i]) + "," + str(bins[i + 1]) + "]")
            names[-1] = names[-1].replace(",inf]",",inf)")
        return names

    def draw_data(self,df,key,itemShowName,describe,*args,**kwargs):#必须指定bin,unit
        per = round(len(df)/self.count_df(df),4) *100

        logging.info("绘制%s %s对应的数据玫瑰分布图，显示名称为%s",describe,key,self.sitename + itemShowName)
        data,dir_data = self.extract_data(df,key)
        plot_data_of_one_item(var=data, dir_a=dir_data, item=self.sitename + itemShowName ,title = describe ,per = str(per) ,*args, **kwargs)

    def draw_all_data(self,*args,**kwargs):
        self.draw_data(df = self.df,describe = "全部",*args,**kwargs)

    def draw_annual_data(self,*args,**kwargs):
        self.annual_df = one_year_df(self.df)
        key = kwargs['key']

        itemShowName = kwargs['itemShowName']
        title = self.sitename + itemShowName + '各向特征值分布'
        plot_distribution_of_df(self.annual_df.groupby('DIR'),key=key,title = title)

        self.draw_data(df = self.annual_df,describe = "全年",*args,**kwargs)

    def draw_month_data(self,*args,**kwargs):
        key = kwargs['key']
        itemShowName = kwargs['itemShowName']

        for m in self.g_m.groups:
            j = self.g_m.get_group(m)
            self.draw_data(df=j, describe=str(m[0]) + '年' + str(m[1]) + "月", *args, **kwargs)
            title = str(m[0]) + '年' + str(m[1]) + "月 " + self.sitename +" " + itemShowName + ' 各向特征值分布'
            plot_distribution_of_df(j.groupby('DIR'),key = key,title=title)

    def draw_season_data(self,*args,**kwargs):
        key = kwargs['key']
        itemShowName = kwargs['itemShowName']

        for s in self.g_s.groups:
            j = self.g_s.get_group(s)
            self.draw_data(df=j, describe=s, *args, **kwargs)
            title = self.sitename + itemShowName + s + "各向特征值分布"
            plot_distribution_of_df(j.groupby('DIR'), key=key, title=title)

    def statics_all_data(self,key,bin,writer = None):
        logging.info("统计全部数据")
        df = self.distribution(df=self.df, key = key,v_bin = bin)
        sheet_name = "全部" + key.replace(r'/', '_')
        if writer:
            df.to_excel(writer, sheet_name=sheet_name , float_format='%.2f')

    def statics_annual_data(self,key,bin,writer = None):
        logging.info("统计全年数据")
        self.annual_df = one_year_df(self.df)
        df = self.distribution(df=self.annual_df,key = key,v_bin = bin)
        sheet_name = "全年" + key.replace(r'/', '_')
        if writer:
            df.to_excel(writer, sheet_name=sheet_name , float_format='%.2f')

    def statics_season_data(self,key,bin,writer = None):
        for s in self.g_s.groups:
            logging.info("统计%s数据",str(s))
            df = self.distribution(df=self.g_s.get_group(s).sort_index(), key = key,v_bin = bin)
            if writer:
                df.to_excel(writer, sheet_name=key.replace(r'/', '_') +  " "+ s, float_format='%.2f')
        if writer:
            gg2 = self.df.groupby(['season','DIR'])
            max = gg2[key].max().unstack()
            max['汇总'] = max.T.max()
            max = max.T
            max['汇总'] = max.T.max()
            max = max.fillna("-")
            max.to_excel(writer,sheet_name = key.replace(r'/', '_') +  " "+ "季最大值",float_format = '%.2f')
            mean = gg2[key].mean().unstack()
            mean['汇总'] = mean.T.mean()
            mean = mean.T
            mean['汇总'] = mean.T.mean()
            mean = mean.fillna("-")
            mean.to_excel(writer,sheet_name = key.replace(r'/', '_') +  " "+ "季平均值",float_format='%.2f')

    def statics_month_data(self,key,bin,writer = None,*args,**kwargs):
        for m in self.g_m.groups:
            logging.info("统计%s数据", str(m[0]) + '年' + str(m[1]) + "月")
            df = self.distribution(df=self.g_m.get_group(m), key = key,v_bin = bin)
            if writer:
                df.to_excel(writer, sheet_name=key.replace(r'/', '_') + " "+ str(m[0]) + '年' + str(m[1]) + "月", float_format='%.2f')
        if writer:
            gg = self.df.groupby(['month','DIR'])
            max = gg[key].max().unstack()
            max['汇总'] = max.T.max()
            max = max.T
            max['汇总'] = max.T.max()
            max = max.fillna("-")
            max.to_excel(writer,sheet_name = key.replace(r'/', '_') +  " "+ "月最大值",float_format = '%.2f')
            mean = gg[key].mean().unstack()
            mean['汇总'] = mean.T.mean()
            mean = mean.T
            mean['汇总'] = mean.T.mean()
            mean  = mean.fillna("-")
            mean.to_excel(writer,sheet_name = key.replace(r'/', '_') +  " "+ "月平均值",float_format='%.2f')

    def statics(self,key,bin,excelfile = None):
        self.max_p.update({key:[]})

        if bin[0] != 0:
            bin.insert(0,0)#见1
        if bin[-1] != inf:
            bin.append(inf)

        if excelfile:
            writer = pd.ExcelWriter(excelfile,engine='xlsxwriter')
        else:
            writer = None
        self.statics_annual_data(key = key,bin = bin, writer = writer)
        self.statics_month_data(key = key,bin = bin, writer = writer)
        self.statics_season_data(key = key,bin = bin, writer = writer)
        logging.info('数据统计结束')
        sheet_names = []
        if writer:
            workbook = writer.book
            format = workbook.add_format()
            format.set_align('center')
            format.set_align('vcenter')
            range1 = "B1:Z18"
            for i in writer.sheets:
                writer.sheets[i].set_column('A:Z', 5, format)
                if '值' not in i:
                    range1 = "B1:K19"
                    sheet_names.append(i)
                writer.sheets[i].conditional_format( range1,{'type': '3_color_scale',
                                             'min_color': "#64be7b",
                                            'mid_color': "#ffeb84",
                                            'max_color': "#f86c6c",})
            writer.close()

            exc = client.gencache.EnsureDispatch("Excel.Application")
            exc.Visible = False
            wb = exc.Workbooks.Open(Filename=excelfile)
            for name in sheet_names:
                sheet = wb.Worksheets(name)
                sheet.Cells(1, 1).Value = "方向/分布"
                sheet.Cells(1, 1).Font.Bold = True
            exc.ActiveWorkbook.Save()
            exc.Quit()

#test = Data({'filename':f},{'key':"Hmean"},{'sitename':"鼠浪湖"},{'itemShowName':"平均波高"},{'bins':h_10_bins},{'unit':"(s)"},{'draw_type':'ALL'},{'workdir':None})
"""
parser = argparse.ArgumentParser(description='绘制玫瑰图程序,请指定文件、绘制列、对应分隔次序等信息')
parser.add_argument('-f', dest='Excelfile', action='store', required=True,help='输入Excel文件')
parser.add_argument('-isNUM', dest='is_EWSN',  action='store_false', default='True', help='方向数据是否为数字格式')
parser.add_argument('-key',dest= 'key',action='store', required=True,help='对应数据列标题')
parser.add_argument('-site',dest = 'sitename', required=True,help='站点名称')
parser.add_argument('-bin',dest='bins', required=True, nargs='*',type = float,help='数据区分间隔')
parser.add_argument('-u',dest='unit', help='数据单位' )
parser.add_argument('-name',dest='name', help='数据列名称' )
parser.add_argument('-mode',dest = 'drawmode',nargs = '*',type = int,help = '1为绘制所有分布图，2为绘制每月分布图，3为绘制每季度分布图')
parser.add_argument('-outdir',dest = 'out',help = '输出文件目录')
parser.add_argument('-log', dest='logging',  action='store_true',default='Fault',help= '日志文件')
parser.add_argument('-Statics', dest='statics',  action='store_true',default='Fault',help= '是否进行数据统计')
parser.add_argument('-Excel', dest='excelfile', help= '输出文件名')
if __name__ == '__main__':
    args = parser.parse_args()
    print(args)
    if args.logging:
        logging.basicConfig(filename='rose_plot.log',level=logging.DEBUG)
    try:
        chdir(args.out)
        logging.info('当前工作路径改变为%s', args.out)
    except:
        logging.error('工作路径有错，请检查',args.out)
    test = Data(filename=args.Excelfile, isEWSN=args.is_EWSN, key=args.key, sitename=args.sitename, bins=args.bins,
                unit='(' + args.unit +')', itemShowName=args.name, draw_type=args.drawmode)
    if args.statics:
        test.statics_to_excel(args.excelfile,args.key,args.bins)
"""
T_bins = [0,4,5,6,7,8,10,12,15]
h_10_bins = [0,0.4,0.6,0.8,1.0,1.2,1.5,2.0,3.0]
f= r"H:\附录C：测波观测月报表（埃及）.xlsx"
chdir(r"H:\埃及")
test = Data(filename=f, isEWSN=False, key='H1/10', sitename="埃及Quseir", bins=h_10_bins,unit='m', itemShowName="1/10波高", draw_type=[])
#test.draw(draw_type=[1,2,3],key = "T1/10",itemShowName = "1/10周期",unit = "s",bins=T_bins)
test.statics(key='H1/10', bin=h_10_bins,excelfile=r"H:\埃及\statics_1_10波高.xlsx")
test.statics(key='T1/10', bin=T_bins,excelfile=r"H:\埃及\statics_1_10周期.xlsx")

print(" *"*20)

#test1 = Data(filename=f, isEWSN=True, key='T1/10', sitename="鼠浪湖", bins=T_bins,unit='(' + 's' +')', itemShowName="1/10周期", draw_type=[1,2,3])
#test2 = Data(filename=f, isEWSN=True, key='Tmean', sitename="鼠浪湖", bins=T_bins,unit='(' + 's' +')', itemShowName="平均周期", draw_type=[1,2,3])


"""
test.draw_all_data(key='T1/f10', sitename="鼠浪湖", bins=T_bins,
                unit='(' + 's' +')', itemShowName="1/10周期")
test.draw_month_data(key='T1/10', sitename="鼠浪湖", bins=T_bins,
                unit='(' + 's' +')', itemShowName="1/10周期")
test.draw_season_data(key='T1/10', sitename="鼠浪湖", bins=T_bins,
                unit='(' + 's' +')', itemShowName="1/10周期")
test.draw_all_data(key='Tmean', sitename="鼠浪湖", bins=T_bins,
                unit='(' + 's' +')', itemShowName="平均周期")
test.draw_month_data(key='Tmean', sitename="鼠浪湖", bins=T_bins,
                unit='(' + 's' +')', itemShowName="平均周期")
test.draw_season_data(key='Tmean', sitename="鼠浪湖", bins=T_bins,
                unit='(' + 's' +')', itemShowName="平均周期")
                """
#test.statics(excelfile=r"H:\鼠浪湖\statics_1_10波高.xlsx",key='H1/10', bin=h_10_bins)
#test.statics(excelfile=r"H:\鼠浪湖\statics_平均周期.xlsx",key='Tmean', bin=T_bins)
#test.statics(excelfile=r"H:\鼠浪湖\statics_1_10周期.xlsx",key='T1/10', bin=T_bins)
