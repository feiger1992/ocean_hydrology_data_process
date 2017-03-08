import sys
import xlsxwriter
import numpy
import datetime
sys.path.append('~/Qt35.py')
import Qt35
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
            'num_format': 'hh:mm',
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
        ###########################################################
        self.data = Qt35.Tide(filename)
        #########################################################################################################################
        x = self.data
        m = lambda t: 0 if t.hour <= 12 else 2

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
            for mon in x.month[site]:

                row1 = count * 42
                count = count + 1  # 第几个月
                ###############################################
                for r in range(row1, 3 + row1):
                    sheet.set_row(r, height=20)
                for r in range(row1 + 3, row1 + 41):
                    sheet.set_row(r, height=14.25)
                o_clock = mon.loc[mon.groupby(mon.index.minute == 0).groups[True], ['tide']]
                ######补全星号
                for ii in range(row1+6,row1+6+o_clock.index[1].daysinmonth):
                    for jj in range(1,35):
                        sheet.write(ii,jj,'*',format_num)
                #######################################################表头
                for i in o_clock.index:
                    t = []
                    h = o_clock.get_value(i, 'tide') * alpha
                    sheet.write(row1 + 5 + i.day, 0, i.date(), format_date)
                    sheet.write(row1 + 5 + i.day, i.hour + 1, h, format_num)
                    t.append(h)  # 当天的
                    o_clock['day'] = o_clock.index.day
                    gg = o_clock['tide'].groupby(o_clock.day)
                    mm = gg.mean()
                    ss = gg.sum()
                    sheet.write(row1 + 5 + i.day, 25, ss.get(i.day) * alpha, format_num)
                    sheet.write(row1 + 5 + i.day, 26, mm.get(i.day) * alpha, format_num)

                    sheet.merge_range(row1 + 6 + i.daysinmonth, 21, row1 + 6 + i.daysinmonth, 24, '月 合 计 ', format_cn)
                    sheet.merge_range(row1 + 7 + i.daysinmonth, 21, row1 + 7 + i.daysinmonth, 24, '月 平 均 ', format_cn)
                    ########底层画线
                for ii in range(21):
                    sheet.write(row1 + 8 + i.daysinmonth, ii, None, format_cn_a_left1)

                ####################################################高低潮位处理
                self.high = mon.loc[mon.groupby(mon.if_max).groups[True], ['tide', 'format_time', 'raising_time']]
                high = self.high
                self.low = mon.loc[mon.groupby(mon.if_min).groups[True], ['tide', 'format_time', 'ebb_time']]
                low = self.low
                t1 = 0
                print('=高潮汇总================================')
                for i in high.dropna().index:
                    sheet.write(row1 + 5 + i.day, 27 + m(i), i, format_time)
                    sheet.write(row1 + 5 + i.day, 28 + m(i), high.loc[i,['tide'] ] * alpha, format_num)
                    t1 += (high.loc[i, 'raising_time']).total_seconds()
                    print(high.loc[i])
                t1 = t1/len(high.dropna().index)
                sheet.write(row1 + 5 + high.tide.idxmax().day, 28 + m(high.tide.idxmax()), high.tide.max() * alpha,
                            format_num_r)
                sheet.write(row1 + 6 + i.daysinmonth, 2, '月 最 高 高 潮 = ' + str(int(high.tide.max() * alpha)),
                            format_cn_a_left)
                sheet.write(row1 + 7 + i.daysinmonth, 2, '潮 时 = ' + high.tide.idxmax().strftime('%H:%M:%S'),
                            format_cn_a_left)

                sheet.write(row1 + 8 + i.daysinmonth, 2, '平均涨潮历时:' + str(int(divmod(t1,3600)[0])) + '小时' + str(int(divmod(t1,3600)[1])) + '分钟',
                            format_cn_a_left1)

                sheet.merge_range(row1 + 8 + i.daysinmonth, 21, row1 + 8 + i.daysinmonth, 25, '月平均高潮潮高=',
                                  format_cn_a_right1)
                sheet.write(row1 + 8 + i.daysinmonth, 26, int(alpha * high.tide.mean()), format_cn_a_left1)
                t2 = 0
                print('=低潮汇总================================')
                for i in low.dropna().index:
                    sheet.write(row1 + 5 + i.day, 31 + m(i), i, format_time)
                    sheet.write(row1 + 5 + i.day, 32 + m(i), low.loc[i, 'tide'] * alpha, format_num)
                    t2 += (low.loc[i, 'ebb_time']).total_seconds()
                    print('=================================')

                    print(low.loc[i])
                    print(low.loc[i, 'ebb_time'].total_seconds())
                print(len(low.index))
                t2 = t2/len(low.dropna().index)

                sheet.write(row1 + 5 + low.tide.idxmin().day, 32 + m(low.tide.idxmin()), low.tide.min() * alpha,
                            format_num_r)
                sheet.write(row1 + 6 + i.daysinmonth, 9, '月 最 低 低 潮 = ' + str(int(low.tide.min() * alpha)),
                            format_cn_a_left)
                sheet.write(row1 + 7 + i.daysinmonth, 9, '潮 时 = ' + low.tide.idxmin().strftime('%H:%M:%S'),
                            format_cn_a_left)

                print(t2)
                sheet.write(row1 + 8 + i.daysinmonth, 9, '平均落潮历时:' +str(int(divmod(t2,3600)[0])) + '小时' + str(int(divmod(t2,3600)[1])) + '分钟',
                            format_cn_a_left1)
                sheet.merge_range(row1 + 8 + i.daysinmonth, 27, row1 + 8 + i.daysinmonth, 32, '月平均低潮潮高=',
                                  format_cn_a_right2)
                sheet.merge_range(row1 + 8 + i.daysinmonth, 33,row1 + 8 + i.daysinmonth,34, int(alpha * low.tide.mean()), format_cn_a_left2)
###############################################################################
                sheet.write(row1 + 6 + i.daysinmonth, 15,
                            '月 平 均 潮 差 = ' + str(int(alpha *numpy.mean(mon['diff'].dropna().values))),
                            format_cn_a_left)
                sheet.write(row1 + 7 + i.daysinmonth, 15, '月 最 大 潮 差 = ' + str(int(mon['diff'].max() * alpha)),
                            format_cn_a_left)
                sheet.write(row1 + 8 + i.daysinmonth, 15, '月 最 小 潮 差 = ' + str(int(mon['diff'].min() * alpha)),
                            format_cn_a_left1)
            #############################################################################
                r_1 = row1+7
                r_2 = r_1+i.daysinmonth-1
                sheet.merge_range(row1+6+ i.daysinmonth,25,row1+6+ i.daysinmonth,26,'=SUM(Z'+str(r_1)+':Z'+str(r_2)+')',format_num)
                sheet.merge_range(row1 + 7 + i.daysinmonth, 25, row1 + 7 + i.daysinmonth, 26,'=AVERAGE(AA' + str(r_1) + ':AA' + str(r_2) + ')',format_num)
                sheet.merge_range(row1 + 6 + i.daysinmonth, 27, row1 + 6 + i.daysinmonth, 28,'=SUM(AC' + str(r_1) + ':AC' + str(r_2) + ')',format_num)
                sheet.merge_range(row1 + 7 + i.daysinmonth, 27, row1 + 7 + i.daysinmonth, 28,'=AVERAGE(AC' + str(r_1) + ':AC' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + i.daysinmonth, 29, row1 + 6 + i.daysinmonth, 30,'=SUM(AE' + str(r_1) + ':AE' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + i.daysinmonth, 29, row1 + 7 + i.daysinmonth, 30,'=AVERAGE(AE' + str(r_1) + ':AE' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + i.daysinmonth, 31, row1 + 6 + i.daysinmonth, 32,'=SUM(AG' + str(r_1) + ':AG' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + i.daysinmonth, 31, row1 + 7 + i.daysinmonth, 32,'=AVERAGE(AG' + str(r_1) + ':AG' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 6 + i.daysinmonth, 33, row1 + 6 + i.daysinmonth, 34,'=SUM(AI' + str(r_1) + ':AI' + str(r_2) + ')', format_num)
                sheet.merge_range(row1 + 7 + i.daysinmonth, 33, row1 + 7 + i.daysinmonth, 34,'=AVERAGE(AI' + str(r_1) + ':AI' + str(r_2) + ')', format_num)

################################################################################标题
                sheet.write(row1, 0, '潮位观测报表:', format_title1)
                sheet.write(row1 + 1, 0, '工程名称：', format_title2)
                sheet.write(row1 + 1, 0, '工程名称：', format_title2)
                sheet.write(row1 + 2, 0, '海区：', format_title2)
                sheet.write(row1 + 1, 16, '高程系统：原始数据高程', format_title2)
                sheet.write(row1 + 2, 16,
                            '观测日期：' + mon.format_time.min().strftime('%Y/%m/%d') +'至'+ mon.format_time.max().strftime('%Y/%m/%d'),format_title2)
                sheet.write(row1 + 1, 29, '测站：' + str(site), format_title2)
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

    def file(self):
        return  self.file
    def high(self):
        return  self.high
tide = Put_tide_in('test2.xlsx','输出.xlsx')
