#encoding = utf-8
from PyQt5.QtWidgets import QFileDialog,QFormLayout,QApplication,QWidget,QMainWindow,QPushButton
from PyQt5.QtGui import QFont
import sys
def split_and_insert(list_of_str,index,length,reverse = True):
    #将位于index位置上的字符串拆分length个字符移入前|后
    str = list_of_str
    if not reverse:
        str.insert(index, str[index][:length])
        str[index+1] = str[index+1][length:]
    else:
        str.insert(index+1,str[index][-length:])
        str[index] = str[index][:-length]
    return str
add_colon_to_time_data=lambda arg :arg[:2]+':'+arg[2:]

def process(line):
    line = line.replace('-', ' ')
    line = ' '.join(line.split())
    datas = line.split(' ')
    datas = list(filter(None, datas))
    if len(datas) == 1:
        return ','.join(datas)+'\n'
    if len(datas[-1:][0]) < 10:
        print(datas[:-1])
        return ','.join(datas)+'\n'
    print(datas[0],'*****',len(datas))
    try:
        if len(datas) == 22:
            split_and_insert(datas,len(datas)-2,3,False)
            split_and_insert(datas, len(datas) - 2, 4)
            split_and_insert(datas, len(datas) - 3, 5)
            split_and_insert(datas, len(datas) - 4, 4)
            split_and_insert(datas, len(datas) - 5, 5)
        else:
            split_and_insert(datas, len(datas) - 2, 4)
            split_and_insert(datas, len(datas) - 4, 4)
        if len(datas[-5]) == 9:
            split_and_insert(datas, len(datas) - 5, 5)
        if len(datas) == 25:
            split_and_insert(datas, len(datas) - 5, 5)
            split_and_insert(datas, len(datas) - 6, 5)

        split_and_insert(datas, len(datas) - 1, 120)
        split_and_insert(datas, 18, 4)
        split_and_insert(datas, 15, 4)
        split_and_insert(datas, 14, 4)
        split_and_insert(datas, 10, 4)
        split_and_insert(datas, 6, 4)
    except:
        return ','.join(datas)+'\n'
    del datas[21]
    for i in [0, 7, 12, 17, 19, 22, 27, 29]:
        datas[i]= add_colon_to_time_data(datas[i])
    return ','.join(datas)+'\n'

def process_file(*args):
    for filename in args:
        filename_to_in = '/'.join(filename.split('/')[:-1]) + '/' + filename.split('/')[-1].replace('.DAT', '.csv')
        try:
            f2 = open(filename_to_in, 'w')
        except PermissionError:
            print('出现错误')
            f2 = open(filename_to_in+'临时文件.csv', 'w')

        line1 = '时间,2min平均风向,2min平均风速,10min平均风向,10min平均风速,最大风速对应风向,最大风速,最大风速出现时间,分钟内最大瞬时风速对应风向,分钟内最大瞬时风速,极大风向,极大风速,极大风速出现时间,分钟降水量,小时累积降水量,气温,最高气温,最高气温出现时间,最低气温,最低气温出现时间,相对湿度,最小相对湿度,最小相对湿度出现时间,水汽压,露点温度,本站气压,最高本站气压,最高本站气压出现时间,最低本站气呀,最低本站气压出现时间,质量控制码,一小时内六十个分钟降水量\n'
        f2.writelines(line1)
        with open(filename) as file:
            for line in file:
                try :
                    line = process(line)
                except IndexError:
                    pass
                f2.writelines(line)
        f2.close()

class m_window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileNames(self, "选取文件", "E:",
                                                   "气象数据文件 (*.DAT);;", options=options)
        process_file(*fileName)

        centralWidget = QWidget()
        self.setWindowTitle('吉布提气象数据处理系统（三航院）')
        self.load_file_Button = QPushButton("数据处理结束,点击结束程序")
        self.load_file_Button.setStyleSheet("""
               QPushButton
               {
                   background-color: rgb(255, 255,255);
                   border:3px solid rgb(170, 170, 170);
               }
               """ )
        self.load_file_Button.setFixedHeight(50)
        self.load_file_Button.setFixedWidth(305)
        font = QFont()
        font.setBold(True)
        font.setPointSize(15)
        self.load_file_Button.setFont(font)
        self.load_file_Button.clicked.connect(self.closeEvent)
        layout = QFormLayout()
        layout.addWidget(self.load_file_Button)
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)

    def closeEvent(self):
        QApplication.instance().quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = m_window()
    ex.show()
    sys.exit(app.exec_())