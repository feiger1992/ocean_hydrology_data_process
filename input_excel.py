import sys
import xlsxwriter
import numpy
import datetime
import pycnnum
sys.path.append('~/Qt35.py')
import Qt35
import lunardate
class Put_tide_in(Qt35.Tide):
    def __init__(self,filename,outfile):
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
            #月平均高潮潮高数字
            #最底部一行的格式
            'font_name': 'Times New Roman',
            'font_size': 7,
            'bottom':1,
            'valign': 'vcenter',
            'align': 'left'
        })
        format_cn_a_left2 = excel_input.add_format({
            #月平均低潮潮高数字
            'font_name': 'Times New Roman',
            'font_size': 7,
            'bottom': 1,
            'right':1,
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
        format_en = excel_input.add_format({
            'border': 1,
            'font_name': 'Times New Roman',
            'font_size': 7,
            'valign': 'vcenter',
            'align': 'center'
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
        format_l= excel_input.add_format({
            'left': 1,
        })
        format_lb= excel_input.add_format({
            'left': 1,
            'bottom':1
        })
        ###########################################################
        self.data = Qt35.Tide(filename)
        #########################################################################################################################
        x = self.data
        #m = lambda t: 0 if t.hour <= 12 else 2
        def check(i,t_all):
            #t_all is a list
            if i.date() in t_all:
                return 2
            else:
                t_all.append(i.date())
                return  0
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
            if max(x.data[site].index)-min(x.data[site].index) < datetime.timedelta(days=50):
                x.month[site] = [((x.month[site][0]).append((x.month[site][1]))).sort_index()]

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
                    for ii in range(row1+6,row1+6+o_clock.index[1].daysinmonth):
                        for jj in range(1,35):
                            if jj == 33 or jj ==31 or jj == 29 or jj == 27:
                                sheet.write(ii,jj,'**',format_num)
                            else:
                                sheet.write(ii, jj, '*', format_num)
                    hangshu = max(mon.index).daysinmonth
                else:
                    hangshu = (max(mon.index).to_pydatetime().date() - min(
                        mon.index).to_pydatetime().date()).days

                #######################################################表头
                for i in o_clock.index:
                    t = []
                    h = o_clock.get_value(i, 'tide') * alpha
                    if len(x.month[site])!= 1:
                        hanghanghang = i.day
                    else:
                        hanghanghang = (i.to_pydatetime().date() - o_clock.index[0].to_pydatetime().date()).days + 1
                    if len(x.month[site]) != 1:
                        sheet.write(row1 + 5 + i.day, 0, i.date(), format_date)
                        sheet.write(row1 + 5 + i.day, 35,self.solar2lunar_d(i), format_cn)
                        sheet.write(row1 + 5 + i.day, i.hour + 1, h, format_num)
                    else:

                        sheet.write(row1 + 5 + hanghanghang, 0, i.date(), format_date)

                        sheet.write(row1 + 5 + hanghanghang, 0, i.date(), format_date)
                        sheet.write(row1 + 5 + hanghanghang, 35,self.solar2lunar_d(i), format_cn)
                        sheet.write(row1 + 5 + hanghanghang, i.hour + 1, h, format_num)
                    t.append(h)  # 当天的
                    o_clock['day'] = o_clock.index.day
                    gg = o_clock['tide'].groupby(o_clock.day)
                    mm = gg.mean()
                    ss = gg.sum()

                    if len(x.month[site]) != 1:
                        sheet.write(row1 + 5 + i.day, 25, ss.get(i.day) * alpha, format_num)
                        sheet.write(row1 + 5 + i.day, 26, mm.get(i.day) * alpha, format_num)
                        sheet.merge_range(row1 + 6 + i.daysinmonth, 21, row1 + 6 + i.daysinmonth, 24, '月 合 计 ', format_cn)
                        sheet.merge_range(row1 + 7 + i.daysinmonth, 21, row1 + 7 + i.daysinmonth, 24, '月 平 均 ', format_cn)
                    else:
                        sheet.write(row1 + 5 + hanghanghang, 25, ss.get(i.day) * alpha, format_num)
                        sheet.write(row1 + 5 + hanghanghang, 26, mm.get(i.day) * alpha, format_num)
                        sheet.merge_range(row1 + 7 + hangshu+1 , 21, row1 + 6 + hangshu, 24, '月 合 计 ', format_cn)
                        sheet.merge_range(row1 + 8 + hangshu +1, 21, row1 + 7 + hangshu, 24, '月 平 均 ', format_cn)

                if len(x.month[site]) == 1:
                    hangshu = hangshu + 1
                            ########底层画线
                for ii in range(21):
                    sheet.write(row1 + 8 + hangshu, ii, None, format_cn_a_left1)
                ####################################################高低潮位处理###################
                self.high = mon.loc[mon.groupby(mon.if_max).groups[True], ['tide', 'format_time', 'raising_time']]
                high = self.high
                self.low = mon.loc[mon.groupby(mon.if_min).groups[True], ['tide', 'format_time', 'ebb_time']]
                low = self.low
                t1 = 0
                t_count = []

                print('=高潮汇总================================')
                for i in high.dropna().index:
                    if len(x.month[site]) == 1:
                        hanghanghang_1 = (i.to_pydatetime().date() - o_clock.index[0].to_pydatetime().date()).days + 1
                    else:
                        hanghanghang_1 = i.day
                    move1= check(i,t_count)  #是否向右移动
                    sheet.write(row1 + 5 + hanghanghang_1, 27 + move1, i, format_time)

                    if i == high.tide.idxmax():
                        sheet.write(row1 + 5 + hanghanghang_1, 28 + move1, high.loc[i, ['tide']] * alpha, format_num_r)
                    else:
                        sheet.write(row1 + 5 + hanghanghang_1, 28 + move1, high.loc[i,['tide'] ] * alpha, format_num)
                    t1 += (high.loc[i, 'raising_time']).total_seconds()
                    print(high.loc[i])
                t1 = t1/len(high.dropna().index)

                sheet.write(row1 + 6 + hangshu, 2, '月 最 高 高 潮 = ' + str(int(high.tide.max() )),
                            format_cn_a_left)

                sheet.write(row1 + 7 + hangshu, 2, '潮 时 = ' + str(high.tide.idxmax().month)+'月'+str(high.tide.idxmax().day)+'日'+str(high.tide.idxmax().hour)+'时'+str(high.tide.idxmax().minute)+'分',
                            format_cn_a_left)

                sheet.write(row1 + 8 + hangshu, 2, '平均涨潮历时:' + str(int(divmod(t1,3600)[0])) + '小时' + str(int(divmod(t1,3600)[1]/60)) + '分钟',
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
                    if i ==low.tide.idxmin():
                        sheet.write(row1 + 5 + hanghanghang_2, 32 + move1, low.loc[i, 'tide'] * alpha, format_num_r)
                    else:
                        sheet.write(row1 + 5 + hanghanghang_2, 32 + move1, low.loc[i, 'tide'] * alpha, format_num)
                    t2 += (low.loc[i, 'ebb_time']).total_seconds()
                    print('=================================')
                    print(low.loc[i])
                t2 = t2/len(low.dropna().index)
                sheet.write(row1 + 6 + hangshu, 9, '月 最 低 低 潮 = ' + str(int(low.tide.min() )),
                            format_cn_a_left)

                sheet.write(row1 + 7 + hangshu, 9, '潮 时 = ' + str(high.tide.idxmin().month)+'月'+str(high.tide.idxmin().day)+'日'+str(high.tide.idxmin().hour)+'时'+str(high.tide.idxmin().minute)+'分',
                            format_cn_a_left)

                print(t2)
                sheet.write(row1 + 8 + hangshu, 9, '平均落潮历时:' +str(int(divmod(t2,3600)[0])) + '小时' + str(int(divmod(t2,3600/60)[1])) + '分钟',
                            format_cn_a_left1)
                sheet.merge_range(row1 + 8 + hangshu, 27, row1 + 8 + hangshu, 32, '月平均低潮潮高=',
                                  format_cn_a_right2)
                sheet.merge_range(row1 + 8 + hangshu, 33,row1 + 8 + hangshu,35, int(alpha * low.tide.mean()), format_cn_a_left2)
###############################################################################
                sheet.write(row1 + 6 + hangshu, 15,
                            '月 平 均 潮 差 = ' + str(int(alpha *numpy.mean(mon['diff'].dropna().values))),
                            format_cn_a_left)
                sheet.write(row1 + 7 + hangshu, 15, '月 最 大 潮 差 = ' + str(int(mon['diff'].max() * alpha)),
                            format_cn_a_left)
                sheet.write(row1 + 8 + hangshu, 15, '月 最 小 潮 差 = ' + str(int(mon['diff'].min() * alpha)),
                            format_cn_a_left1)
            #############################################################################
                r_1 = row1+7
                r_2 = r_1+hangshu-1
                sheet.merge_range(row1+6+ hangshu,25,row1+6+ hangshu,26,'=SUM(Z'+str(r_1)+':Z'+str(r_2)+')',format_num)
                sheet.merge_range(row1 + 7 + hangshu, 25, row1 + 7 + hangshu, 26,'=AVERAGE(AA' + str(r_1) + ':AA' + str(r_2) + ')',format_num)
                sheet.merge_range(row1 + 6 + hangshu, 27, row1 + 6 + hangshu, 28,'=SUM(AC' + str(r_1) + ':AC' + str(r_2) + ')',format_num)
                sheet.merge_range(row1 + 7 + hangshu, 27, row1 + 7 + hangshu, 28,'=AVERAGE(AC' + str(r_1) + ':AC' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 29, row1 + 6 + hangshu, 30,'=SUM(AE' + str(r_1) + ':AE' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 29, row1 + 7 + hangshu, 30,'=AVERAGE(AE' + str(r_1) + ':AE' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 31, row1 + 6 + hangshu, 32,'=SUM(AG' + str(r_1) + ':AG' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 31, row1 + 7 + hangshu, 32,'=AVERAGE(AG' + str(r_1) + ':AG' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + hangshu, 33, row1 + 6 + hangshu, 34,'=SUM(AI' + str(r_1) + ':AI' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + hangshu, 33, row1 + 7 + hangshu, 34,'=AVERAGE(AI' + str(r_1) + ':AI' + str(r_2) + ')', format_num)
############################
                sheet.write_blank(row1 + 6 + hangshu, 35,None, format_num)
                sheet.write_blank(row1 + 7 + hangshu, 35,None, format_num)
#######
                sheet.write_blank(row1 + 6 + hangshu,0,None,format_l)
                sheet.write_blank(row1 + 7 + hangshu,0,None,format_l)
                sheet.write_blank(row1 + 8 + hangshu,0,None,format_lb)

################################################################################标题
                sheet.write(row1, 0, '潮位观测报表:', format_title1)
                sheet.write(row1 + 1, 0, '工程名称：', format_title2)
                sheet.write(row1 + 1, 0, '工程名称：', format_title2)
                sheet.write(row1 + 2, 0, '海区：', format_title2)
                sheet.write(row1 + 1, 16, '高程系统：原始数据高程', format_title2)
                sheet.write(row1 + 2, 16,
                            '观测日期：' + mon.format_time.min().strftime('%Y/%m/%d') +'至'+ mon.format_time.max().strftime('%Y/%m/%d'),format_title2)
                sheet.write(row1 + 1, 29, '测站：' + str(site).replace('原始数据',''), format_title2)
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
                print('正在写入'+site+'站第'+str(count)+'个月')
            print(str(site) +'写入结束')

        print(self.file+'文件写入完成')
        excel_input.close()
    def data(self):
        return self.data

    def solar2lunar_d(self,d):
        #将阳历转化为阴历
        #汉字格式改写
        han_zi =  pycnnum.num2cn(lunardate.LunarDate.fromSolarDate(d.year, d.month, d.day).day)
        han_zi = han_zi.replace('一十','十')
        if len(han_zi) == 3:
            return han_zi.replace('二十','廿')
        if len(han_zi) == 1:
            return '初'+han_zi
        else:
            return han_zi
    def file(self):
        return  self.file
    def high(self):
        return  self.high
tide = Put_tide_in("E:\★★★★★CODE★★★★★\程序调试对比\潮汐模块\对比潮位特征值（东水港村）\YSW 2012-03计算.xls","E:\★★★★★CODE★★★★★\程序调试对比\潮汐模块\对比潮位特征值（东水港村）\潮位报表2012-3（程序结果）.xls")


