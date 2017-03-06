import  sys
import xlsxwriter
sys.path.append('~/Qt35.py')
import Qt35
chaoshi = "潮时"
chaogao = "潮时"

x = Qt35.Tide('test.xlsx')
excel_input = xlsxwriter.Workbook('潮位报表.xlsx')

########################################
format_title1=excel_input.add_format({
    'bold':1,
    'font_name':'宋体',
    'font_size':14,
    'valign':'vcenter'
})
format_title2 = excel_input.add_format({
    'bold': True,
    'font_name':'宋体',
    'font_size':12,
    'valign': 'vcenter'
})
format_cn = excel_input.add_format({
    'border':1,
    'font_name': '宋体',
    'font_size': 8,
    'valign': 'vcenter',
    'align':'center'
})
format_cn_no = excel_input.add_format({
    'left':1,
    'right':1,
    'font_name': '宋体',
    'font_size': 8,
    'valign': 'vcenter',
    'align':'center'
})
format_cn_l = excel_input.add_format({
    'bottom':1,
    'left':1,
    'right':1,
    'font_name': '宋体',
    'font_size': 8,
    'valign': 'vcenter',
    'align':'left'
})
format_cn_r = excel_input.add_format({
    'top':1,
    'left':1,
    'right':1,
    'font_name': '宋体',
    'font_size': 8,
    'valign': 'vcenter',
    'align':'right'
})
format_en = excel_input.add_format({
    'border': 1,
    'font_name': 'Times New Roman',
    'font_size': 7,
    'valign': 'vcenter',
    'align': 'center'
})
format_date = excel_input.add_format({
    'border':1,
    'num_format':'mm/dd',
    'font_name':'Times New Roman',
    'font_size': 7,
    'valign':'vcenter',
    'align':'center',
})
##################################

def title():
    sheet.write(row1, 0, '潮位观测报表', format_title1)
    sheet.write(row1 + 1, 0, '工程名称：', format_title2)
    sheet.write(row1 + 1, 0, '工程名称：', format_title2)
    sheet.write(row1 + 2, 0, '海区：', format_title2)
    sheet.write(row1 + 1, 16, '高程系统：原始数据高程', format_title2)
    sheet.write(row1 + 2, 16,
                '观测日期：' + mon.format_time.min().strftime('%Y/%m/%d') + mon.format_time.max().strftime('%Y/%m/%d'),
                format_title2)
    sheet.write(row1 + 1, 29, '测站：' + str(site), format_title2)
    sheet.write(row1 + 2, 29, '单位：cm', format_title2)

    sheet.write(row1 + 3, 0, chaoshi, format_cn_r)
    sheet.write(row1 + 4, 0, chaogao, format_cn_no)
    sheet.write(row1 + 5, 0, '日期', format_cn_l)
    for c in range(1, 25):
        sheet.merge_range(3, c, 5, c, c - 1, format_cn)
    sheet.merge_range(row1 + 3, 25,row1+5,25, '合计',format_cn)
    sheet.merge_range(row1 + 3, 26,row1+5,26, '平均',format_cn)
    sheet.merge_range(row1 + 3, 27,row1+3,28, '高潮',format_cn)

    sheet.write(row1 + 4, 27, chaoshi,format_cn)
    sheet.write(row1 + 5, 27, '', format_cn)
    sheet.write(row1 + 4, 28, chaogao,format_cn)
    sheet.write(row1 + 5, 28, '', format_cn)
    sheet.merge_range(row1 + 3, 29,row1+3,30, '高潮',format_cn)

    sheet.write(row1 + 4, 29, chaoshi,format_cn)
    sheet.write(row1 + 5, 29, '', format_cn)
    sheet.write(row1 + 4, 30, chaogao,format_cn)
    sheet.write(row1 + 5, 30, '', format_cn)
    sheet.merge_range(row1 + 3, 31,row1+3,32, '低潮',format_cn)

    sheet.write(row1 + 4, 31, chaoshi,format_cn)
    sheet.write(row1 + 5, 31, '', format_cn)
    sheet.write(row1 + 4, 32, chaogao,format_cn)
    sheet.write(row1 + 5, 32, '', format_cn)

    sheet.merge_range(row1 + 3, 33,row1+3,34, '低潮',format_cn)

    sheet.write(row1 + 4, 33, chaoshi,format_cn)
    sheet.write(row1 + 5, 33, '', format_cn)
    sheet.write(row1 + 4, 34, chaogao,format_cn)
    sheet.write(row1 + 5, 34, '', format_cn)

    sheet.merge_range(row1 + 3, 35,row1+5,35, '日期',format_cn)

for site in x.data.keys():

    sheet = excel_input.add_worksheet(name=site)
    count = 0
    ############################
    sheet.set_column('A:A',7)
    sheet.set_column('B:Y',2.75)
    sheet.set_column('Z:AI', 3.5)
    sheet.set_column('AJ:AJ', 5)
    ##############################
    for mon in x.month[site]:
        row1 = count*41
        title()
        count = count + 1  # 第几个月
        ###############################################
        for r in range(row1,3+row1):
            sheet.set_row(r,height=20)
        for r in range (row1+3,row1+39):
            sheet.set_row(r,height=14.25)
        #######################################################表头
        for i in mon.format_time:
            sheet.write(row1+5+i.day,0,i.date(),format_date)
######################################
excel_input.close()

def o_clock(x):
    if x.minute == 0:
        return
    else:
        pass
