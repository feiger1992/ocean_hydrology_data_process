from PyQt5.QtCore import QDir, Qt
from PyQt5.QtCore import QModelIndex, Qt,QDate,QUrl,pyqtSignal,QObject
from PyQt5.QtGui import QStandardItemModel,QValidator,QDoubleValidator
from PyQt5.QtWidgets import QApplication, QMainWindow,QAction,QLabel,QVBoxLayout, QFormLayout,QLineEdit,QSpinBox, QTableView,QFileDialog,QTextBrowser,QLayout,QGroupBox,QHBoxLayout,QTabWidget,QPushButton,QWidget,QCheckBox,QMessageBox,QDialog,QDateTimeEdit,QComboBox,QDialogButtonBox,QErrorMessage,QDoubleSpinBox,QInputDialog,QTextEdit
from PyQt5.QtWebKitWidgets  import QWebView
from threading import Thread
import sys,time
sys.path.append('~/tide.py')
sys.path.append('~/current.py')
from tide import Process_Tide
from current import Current

dir_of_file = lambda filename: filename.replace(list(filename.split(sep='\\'))[-1], '')


# 通知栏         QMessageBox.information(self,'通知标题',"内容")
class M_window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.show_content={}
        self.initUI()
        global sig
        sig = All_Signal_event()
        self.set_signal_event()

    def initUI(self):

        self.setWindowTitle('')
        self.resize(1000, 618)
        self.statusBar()
        #self.createMenu()
        centralWidget = QWidget()

        self.shower = QTabWidget()
        self.allGroup = QGroupBox()
        self.mid_l_Group = QGroupBox()
        self.mid_s_Group = QGroupBox("测验概况")
        self.Tabs = QTabWidget()
        self.Generate_all_button = QPushButton("生成报告及文档")
        centralLayout =  QVBoxLayout()
        centralLayout.addWidget(self.shower, stretch=40)
        centralLayout.addWidget(self.allGroup, stretch=10)
        centralWidget.setLayout(centralLayout)
        self.setCentralWidget(centralWidget)

        allGroupLayout = QHBoxLayout()
        allGroupLayout.addWidget(self.mid_l_Group, stretch=3)
        allGroupLayout.addWidget(self.Tabs, stretch=30)
        self.allGroup.setLayout(allGroupLayout)
        #测试潮流，将潮位部分取消
        self.tides_control = TideTab()
        self.Tabs.addTab(self.tides_control,"潮位")

        self.current_control = Current_Tab()
        self.Tabs.addTab(self.current_control, '潮流')

        mid_l_GroupLayout = QVBoxLayout()
        mid_l_GroupLayout.addWidget(self.mid_s_Group)
        mid_l_GroupLayout.addWidget(self.Generate_all_button)
        self.mid_l_Group.setLayout(mid_l_GroupLayout)

        mid_s_GroupLayout = QFormLayout()
        Project_Name = QLineEdit()
        Sea_Area = QLineEdit()
        Button_Save_Info = QPushButton("确定")
        mid_s_GroupLayout.addRow(QLabel("工程名称"), Project_Name)
        mid_s_GroupLayout.addRow(QLabel("测验海域"), Sea_Area)
        mid_s_GroupLayout.addRow(Button_Save_Info)
        self.mid_s_Group.setLayout(mid_s_GroupLayout)

    def set_signal_event(self):
        sig.load_tide_file_done.connect(self.tide_site_element_event)
        sig.msg_to_show.connect(self.show_msg_statusbar)

    def show_msg_statusbar(self,msg):
        self.statusBar().showMessage(msg)

    def createMenu(self):
        tide_load = QAction('载入数据', self)
        tide_load.setStatusTip('从Excel文件读入潮位数据，不同站点位于不同sheet,列标题应为Tide,Time')
#        tide_load.triggered.connect(self.tide)

        generate_tide_sheet = QAction('&保存潮位报表', self)
        generate_tide_sheet.setStatusTip('保存潮位月报表')
 #       generate_tide_sheet.triggered.connect(self.save_tide)

        tide_statics = QAction('&各月潮位特征值汇总', self)
        tide_statics.setStatusTip('统计各月潮位极值,涨落潮历时等')
 #       tide_statics.triggered.connect(self.out_put_data)

        tide_harmonic_analysis = QAction('&调和分析', self)
        tide_harmonic_analysis.setStatusTip('对潮位数据进行调和分析，计算各分潮特征值')
  #      tide_harmonic_analysis.triggered.connect(self.harmonic_analysis)

        load_stream_E_N = QAction('&载入数据(格式为E-N)', self)
        load_stream_E_N.setStatusTip('对应sheet名应为‘E’,‘N’')

        load_stream_V_D = QAction('&载入数据(格式为V-D)', self)
        load_stream_V_D.setStatusTip('对应sheet名应为‘V’,‘D’')

        get_button = QAction('&去除底部干扰数据', self)
        get_button.setStatusTip('找底')

        constant_3 = QAction('&常数法（三点）', self)
        constant_6 = QAction('&常数法（六点）', self)

        power_3 = QAction('&幂函数法（三点）', self)
        power_6 = QAction('&幂函数法（六点）', self)

        stream_statics = QAction('&统计潮流特征', self)
        stream_statics.setStatusTip('统计最大涨落潮流速、涨落潮历时等特征数据并输出Excel')

        generate_stream_sheet = QAction('&生成潮流报表', self)

        menubar = self.menuBar()
        TIDE = menubar.addMenu('&潮位')
        TIDE.addAction(tide_load)

        TIDE.addAction(generate_tide_sheet)
        TIDE.addAction(tide_statics)
        TIDE.addAction(tide_harmonic_analysis)

        STREAM = menubar.addMenu('&潮流')
        LOAD_STREAM = STREAM.addMenu('&载入潮流数据')
        LOAD_STREAM.addAction(load_stream_E_N)
        LOAD_STREAM.addAction(load_stream_V_D)

        STREAM.addSeparator()

        STREAM.addAction(get_button)

        STREAM.addSeparator()

        FENCENG = STREAM.addMenu('&分层')
        FENCENG.addAction(constant_6)
        FENCENG.addAction(power_6)
        FENCENG.addSeparator()
        FENCENG.addAction(constant_3)
        FENCENG.addAction(power_3)

        STREAM.addSeparator()
        STREAM.addAction(stream_statics)
        STREAM.addAction(generate_stream_sheet)
        # 整体布局

    #def createTideTab(self):
    #    self.tideTab = QTabWidget

    def display(self, category, title, content, tip=None):
        if category =='HTML':
            self.show_content.update({title+'(web)':QWebView()})
            self.shower.addTab(self.show_content[title+'(web)'],title+'(web)')
            self.show_content[title+'(web)'].setHtml(content)
            self.shower.setCurrentWidget(self.show_content[title+'(web)'])
            if tip:
                self.shower.setTabToolTip(self.shower.count() - 1, tip)
        if category =='TXT':
            self.show_content.update({title:QTextEdit()})
            self.shower.addTab(self.show_content[title],title)
            #self.show_content[title].setHtml(content)#试试效果
            self.show_content[title].setReadOnly(True)
            self.show_content[title].setPlainText(content)
            self.shower.setCurrentWidget(self.show_content[title])
            if tip:
                self.shower.setTabToolTip(self.shower.count() - 1, tip)

    def display_tide_HTML(self, filename, site, tip):
        self.display(category='HTML', title=site, content=t[filename].html[site], tip=tip)

    def display_tide_TXT(self, filename, site, tip):
        columns = [ 'tide', 'tide_init', 'if_min', 'if_max', 'diff', 'raising_time', 'ebb_time']
        self.display(category='TXT', title=site,
                     content=t[filename].data[site].to_string(columns=columns, index_names=False).replace('None',
                                                                                                          ' ' * 4).replace(
                         'False', ' ' * 5), tip=tip)

    def display_tide_harmonic_result(self, filename, site, tip):
        self.display(category='TXT', title=site + '调和分析结果', content=t[filename].harmonic_result[site].classic_style(),
                     tip=tip)

    def tide_process(self, filename, sitename, threshold):
        preprocess_thread = Thread(target=t[filename].preprocess, kwargs={'s': sitename, 'threshold': threshold})
        preprocess_thread.start()
        second = 0
        while preprocess_thread.is_alive():
            sig.msg_to_show.emit('正在对' + sitename + '站点数据进行预处理，用时' + str(second) + '秒')
            time.sleep(0.1)
            second += 0.1
        sig.msg_to_show.emit(sitename + '站数据预处理结束，请继续操作')

    def tide_site_element_event(self,file,site):
        self.tides_control.sites_control[site].signal.tide_preprocess.connect(self.tide_process)
        self.tides_control.sites_control[site].signal.show_tide_HTML_signal.connect(self.display_tide_HTML)
        self.tides_control.sites_control[site].signal.show_tide_TXT_signal.connect(self.display_tide_TXT)
        self.tides_control.sites_control[site].signal.show_tide_harmonic_result_signal.connect(
            self.display_tide_harmonic_result)


class TideTab(QWidget):
    def __init__(self,parent = None):
        super(TideTab,self).__init__(parent)
        global t
        t = {}
        self.layout_element()

    def layout_element(self):

        self.tide_l_Group = QGroupBox("潮位文件")
        self.tide_file_button = QPushButton("导入潮位文件",clicked=self.open_file)
        self.auto_Preprocess = QCheckBox("打开文件时进行预处理")
        self.if_one_sheet = QCheckBox("仅载入首个工作表")
        self.check_multifile = QCheckBox("多个文件")
        self.check_multifile.setChecked(True)
        self.auto_Preprocess.setChecked(True)
        self.if_one_sheet.setChecked(True)
        #self.check_if_null = QCheckBox("潮位补全")
        self.del_data = QPushButton("删除当前潮位数据",clicked = self.del_all_data)
        self.tide_d_Group = QGroupBox("其他")

        self.errorMessageDialog = QErrorMessage(self)
        self.errorLabel = QLabel()
        self.errorButton = QPushButton("确定")

        tide_l_Group_Layout = QVBoxLayout()
        tide_l_Group_Layout.addWidget(self.tide_file_button)
        tide_l_Group_Layout.addWidget(self.auto_Preprocess)
        tide_l_Group_Layout.addWidget(self.if_one_sheet)
        tide_l_Group_Layout.addWidget(self.tide_d_Group)

        self.tide_l_Group.setLayout(tide_l_Group_Layout)

        tide_d_Group_Layout = QVBoxLayout()

        tide_d_Group_Layout.addWidget(self.check_multifile)
        tide_d_Group_Layout.addWidget(self.del_data)

        self.tide_d_Group.setLayout(tide_d_Group_Layout)
        self.sites_Tab = QTabWidget()

        #调试site_tab 用
        #self.sites_Tab.addTab(Tide_Site_Tab(), '未载入潮位数据')
        #self.welcome_init_tab()

        mainLayout = QHBoxLayout()
        mainLayout.addWidget(self.tide_l_Group)
        mainLayout.addWidget(self.sites_Tab)
        self.setLayout(mainLayout)

        self.sites_control= {}

    def welcome_init_tab(self):
        self.init_tab = Welcome_Tab("请载入潮位数据文件")
        self.sites_Tab.addTab(self.init_tab, '未载入潮位数据')
        self.init_tab.load_file_button.clicked.connect(self.open_file)

    def open_file(self):
        options = QFileDialog.Options()
        if self.check_multifile.isChecked():
            fileName,_ = QFileDialog.getOpenFileNames(self,"选取文件", "E:",
                                           "97-03表格文档 (*.xls);;Excel文件(*.xlsx)", options=options)
        else:
            fileName, filetype = QFileDialog.getOpenFileName(self, "选取文件", "E:", "97-03表格文档 (*.xls);;Excel文件(*.xlsx)",
                                                             options=options)
            fileName=[fileName]
        #测试用
        #fileName = ["E:\★★★★★CODE★★★★★\Git\codes\ocean_hydrology_data_process\测试潮位.xls"]

        for f in fileName:
            try:
                t.update({f:Process_Tide(f,only_first_sheet = self.if_one_sheet.isChecked())})

            except:
                self.errorMessage(EOFError)
            if self.auto_Preprocess.isChecked():
                for i in t.values():
                    for sitename in i.sites:
                        tide_thread = Thread(target=i.preprocess,kwargs={'s':sitename})
                        second = 0
                        tide_thread.start()
                        while tide_thread.is_alive():
                            sig.msg_to_show.emit("正在进行数据预处理，数据来自文件" + fileName[0] + ',处理用时' + str(second) + '秒')
                            time.sleep(0.1)
                            second += 0.1
                        sig.msg_to_show.emit(fileName[0] + '文件处理完毕')
                sig.msg_to_show.emit("文件全部处理完毕，请继续操作")

        for filename,process_tide in t.items():
            #i为文件名
            #j为process calss
            for site in process_tide.sites:
                self.sites_control.update({site: Tide_Site_Tab(site,filename)})
                self.sites_Tab.addTab(self.sites_control[site],site)
                self.sites_Tab.setTabToolTip(self.sites_Tab.count()-1,filename)

                sig.load_tide_file_done.emit(filename, site)#信号绑定
                self.sites_control[site].signal_connect()
                self.sites_control[site].Save_Sheet_Button.clicked.connect(self.sites_control[site].save_one_site_tide_clicked)

        if t and  self.sites_Tab.tabText(0) == '未载入潮位数据':
            self.sites_Tab.removeTab(0)

    def del_all_data(self):
        MESSAGE = "将删除所有已载入的潮位数据，是否确定？"
        reply = QMessageBox.question(self, "删除潮位数据",
                                     MESSAGE,
                                     QMessageBox.Yes | QMessageBox.No )
        if reply == QMessageBox.Yes:
            t.clear()
            self.clear_tab()
            self.welcome_init_tab()
        else:
            pass
        sig.msg_to_show.emit("所载入数据已全部删除")

    def clear_tab(self):
        while self.sites_Tab.count() !=0 :
            self.sites_Tab.removeTab(0)

    def errorMessage(self,Wrong_msg):
        self.errorMessageDialog.showMessage(Wrong_msg)
        self.errorLabel.setText("确定")

class Site_Tide_signal_event(QObject):
    tide_preprocess = pyqtSignal(str, str, float)
    show_tide_HTML_signal = pyqtSignal(str, str, str)
    show_tide_TXT_signal = pyqtSignal(str, str, str)
    save_one_site_signal = pyqtSignal(str, str)
    show_tide_harmonic_result_signal = pyqtSignal(str, str, str)

class All_Signal_event(QObject):
    load_tide_file_done = pyqtSignal(str, str)
    msg_to_show = pyqtSignal(str)
class Tide_Site_Tab(QWidget):
    def __init__(self,sitename,filename,parent = None):
        super(Tide_Site_Tab, self).__init__(parent)
        self.layout_element()
        self.sitename = sitename
        self.filename = filename
        self.signal = Site_Tide_signal_event()

    def layout_element(self):
        self.lu_Group = QGroupBox("测点位置")
        self.ld_Group = QGroupBox("起止时间")
        self.mid_Group_out = QGroupBox("潮位查看")
        self.mid_Group_in = QGroupBox("调整去噪度")
        self.r_Group = QGroupBox("结果输出")
        self.r_in_Group_u = QGroupBox("调和分析")
        self.r_in_Group_d = QGroupBox("文档输出")
        #self.add_Group1 = QGroupBox("施测信息")
        self.add_Group2 = QGroupBox("高程基准")
        self.add_Group = QGroupBox()

        """add_Group1 = QFormLayout()
        self.Survey_man = QLineEdit()
        self.Survey_Implement = QComboBox()
        self.Survey_Implement.addItem("Tide Master")
        self.Survey_Implement.addItem("其他设备（待补充）")
        self.Survey_Implement.addItem("自定义")
        self.Survey_Implement_Series_num = QLineEdit()
        add_Group1.addRow(QLabel("测量人员"),self.Survey_man)
        add_Group1.addRow(QLabel("采用仪器"), self.Survey_Implement)
        add_Group1.addRow(QLabel("仪器编号"),self.Survey_Implement_Series_num)
        self.add_Group1.setLayout(add_Group1)"""

        add_Group2 = QFormLayout()
        self.Altitudinal_System = Altitudinal_Select()
        self.Altitudinal_Add = QDoubleSpinBox()
        self.Altitudinal_Add.setRange(-100,100)
        self.Altitudinal_Add.setToolTip("单位与原数据一致，计算方式为加")
        self.confirm_change_altitudinal = QPushButton('改变数据高程',clicked=self.change_altitude)
        self.confirm_change_altitudinal.setToolTip("单位与原数据一致，计算方式为加")
        add_Group2.addRow(QLabel("数据基准面"),self.Altitudinal_System.combobox )
        add_Group2.addRow(QLabel("改变基准面"), self.Altitudinal_Add)
        add_Group2.addRow(self.confirm_change_altitudinal)
        self.add_Group2.setLayout(add_Group2)

        add_Group = QHBoxLayout()
        #add_Group.addWidget(self.add_Group1)
        add_Group.addWidget(self.add_Group2)
        self.add_Group.setLayout(add_Group)

        lu_GroupLayout = QFormLayout()
        self.onlyDouble = QDoubleValidator()
        self.Longitude = QLineEdit()
        self.Latitude = QLineEdit()
        self.Longitude.setValidator(self.onlyDouble)
        self.Latitude.setValidator(self.onlyDouble)

        lu_GroupLayout.addRow(QLabel("经度"))
        lu_GroupLayout.addRow(self.Longitude)
        lu_GroupLayout.addRow(QLabel("纬度"))
        lu_GroupLayout.addRow(self.Latitude)
        self.lu_Group.setLayout(lu_GroupLayout)

        ld_GroupLayout = QFormLayout()
        self.start_time = QDateTimeEdit()
        self.end_time = QDateTimeEdit(QDate.currentDate())
        ld_GroupLayout.addRow(QLabel("开始时间"))
        ld_GroupLayout.addRow(self.start_time)
        ld_GroupLayout.addRow(QLabel("结束时间"))
        ld_GroupLayout.addRow(self.end_time)
        self.ld_Group.setLayout(ld_GroupLayout)

        mid_Group_out = QVBoxLayout()
        self.Show_tide_data_button1 = QPushButton("&潮位查看\n（WEB模式）",clicked=self.show_tide_HTML_click)

        self.Show_tide_data_button2 = QPushButton("&潮位查看\n（普通模式）",clicked = self.show_tide_TXT_click)
        self.Process_method_select = QComboBox()
        self.Process_method_select.addItem("小波分析")
        self.Process_method_select.addItem("数值平滑")
        self.confirmButton = QPushButton("数据处理",clicked=self.preprocess_click)
        self.Process_threshold = QSpinBox()
        self.Process_threshold.setRange(1, 100)
        self.Process_threshold.setValue(5)
        mid_Group_out.addWidget(self.Show_tide_data_button1)
        mid_Group_out.addWidget(self.Show_tide_data_button2)
        mid_Group_out.addSpacing(20)
        #mid_Group_out.addWidget(self.mid_Group_in,stretch=4)
        mid_Group_in = QFormLayout()
        mid_Group_in.addRow(self.Process_method_select)
        mid_Group_in.addRow("阀值",self.Process_threshold)
        mid_Group_in.addRow(self.confirmButton)
        self.mid_Group_out.setLayout(mid_Group_out)
        self.mid_Group_in.setLayout(mid_Group_in)

        r_Group = QFormLayout()
        self.Save_Sheet_Button = QPushButton("保存潮位报表")
        self.Save_Mid_Data_Button = QPushButton("保存中间结果", clicked=self.save_one_site_mid_data)
        self.Astronomical_Tide_Analysis = QPushButton("天文潮计算", clicked=self.Astronomical_analysis)
        self.analysis_by_init_data = QCheckBox("使用原始数据计算")
        self.Theoretical_Mininum_tidal_leval = QPushButton("计算对应理论最低潮面")
        self.Theoretical_Mininum_tidal_leval.setEnabled(False)
        self.Is_Output_Report = QCheckBox("输出文字报告")
        self.Is_Output_mid_Document = QCheckBox("输出分析文档")
        r_Group.addRow(self.Save_Sheet_Button,self.Save_Mid_Data_Button)
        r_Group.addRow(self.r_in_Group_u)
        r_in_Group_u = QFormLayout()
        r_in_Group_u.addRow(self.Astronomical_Tide_Analysis)
        r_in_Group_u.addRow(self.analysis_by_init_data)
        r_in_Group_u.addRow(self.Theoretical_Mininum_tidal_leval)

        self.r_in_Group_u.setLayout(r_in_Group_u)
        r_Group.addRow(self.r_in_Group_d)
        r_in_Group_d = QFormLayout()
        r_in_Group_d.addRow(self.Is_Output_Report)
        r_in_Group_d.addRow(self.Is_Output_mid_Document)
        self.r_in_Group_d.setLayout(r_in_Group_d)
        self.r_Group.setLayout(r_Group)

        mainLayout = QHBoxLayout()
        mainLayout.addWidget(self.lu_Group)
        mainLayout.addWidget(self.ld_Group)
        mainLayout.addWidget(self.add_Group)
        mainLayout.addWidget(self.mid_Group_out)
        mainLayout.addWidget(self.mid_Group_in)
        mainLayout.addWidget(self.r_Group)
        self.setLayout(mainLayout)

    def preprocess_click(self):
        self.signal.tide_preprocess.emit(self.filename, self.sitename, self.Process_threshold.value())

    def show_tide_HTML_click(self):
        self.generate_tide_HTML()
        self.signal.show_tide_HTML_signal.emit(self.filename, self.sitename,
                                               '处理阀值为：' + str(self.Process_threshold.value()))

    def show_tide_TXT_click(self):
        self.signal.show_tide_TXT_signal.emit(self.filename, self.sitename,
                                              '处理阀值为：' + str(self.Process_threshold.value()))

    def signal_connect(self):
        self.signal.save_one_site_signal.connect(self.save_one_site_tide)

    def generate_tide_HTML(self):
        html_process_thread = Thread(target=t[self.filename].display,kwargs={'site':self.sitename})
        html_process_thread.start()
        second = 0
        while html_process_thread.is_alive():
            sig.msg_to_show.emit('正在生成潮位对应HTML文件，处理用时' + str(second) + '秒')
            time.sleep(0.2)
            second += 0.2
        sig.msg_to_show.emit(self.filename + '文件中的' + self.sitename + '站HTML潮位显示处理完毕')

    def save_one_site_tide_clicked(self):
        self.signal.save_one_site_signal.emit(self.filename,self.sitename)

    def save_something_to_excel(self,fun,args,file_type_in_chinese):

        options = QFileDialog.Options()
        savefileName, _ = QFileDialog.getSaveFileName(self, "保存"+file_type_in_chinese, "E:",
                                                      "Excel文件 (*.xlsx)", options=options)
        args.update({'outfile': savefileName})
        save_tide_thread = Thread(target=fun,
                                  kwargs=args)
        second = 0
        save_tide_thread.start()
        while save_tide_thread.is_alive():
            sig.msg_to_show.emit('正在生成' + self.sitename + file_type_in_chinese + '，已用时' + str(second) + '秒，请稍等')
            time.sleep(0.1)
            second += 0.1
        sig.msg_to_show.emit(self.sitename + file_type_in_chinese + '文件生成结束')

    def save_one_site_mid_data(self):
        self.save_something_to_excel(t[self.filename].out_put_mid_data, {'sitename': self.sitename},'计算过程文件')

    def save_one_site_tide(self):
        self.save_something_to_excel(t[self.filename].output,{ 'sitename': self.sitename,
                                          'time_start': self.start_time.dateTime().toPyDateTime()
                                      , 'time_end': self.end_time.dateTime().toPyDateTime()
                                      , 'reference_system': self.Altitudinal_System.combobox.currentText()},'潮位报表')

    def change_altitude(self):
        t[self.filename].change_altitude(self.sitename,diff=self.Altitudinal_Add.value())

        sig.msg_to_show.emit('高程处理结束')

    # def tideanalysis(self,sitename),:

    def Astronomical_analysis(self):
        harmonic_thread = Thread(target=t[self.filename].harmonic_analysis,
                                 kwargs={'site': self.sitename, 'if_init': self.analysis_by_init_data.isChecked()})
        harmonic_thread.start()
        second = 0
        while harmonic_thread.is_alive():
            sig.msg_to_show.emit('正在进行调和分析，处理用时' + str(second) + '秒')
            time.sleep(0.2)
            second += 0.2
        sig.msg_to_show.emit(self.filename + '文件中的' + self.sitename + '站调和分析处理完毕')
        if self.analysis_by_init_data.isChecked():
            s = '原始数据'
        else:
            s = '去噪后的数据'
        self.signal.show_tide_harmonic_result_signal.emit(self.filename, self.sitename, s)
        self.Theoretical_Mininum_tidal_leval.setEnabled(True)

class Altitudinal_Select(QDialog):
    def __init__(self,parent = None):
        super(Altitudinal_Select,self).__init__(parent)
        self.combobox = QComboBox()
        self.combobox.addItem("当地理论最低潮面")
        self.combobox.addItem("平均海平面")
        self.combobox.addItem("85高程")
        self.combobox.addItem("自定义")
        self.combobox.activated.connect(self.changed)
    def changed(self,index):
        if index == 0:
            self.select = "当地理论最低潮面"
        elif index ==1:
            self.select = "平均海平面"
        elif index ==2:
            self.select = "85高程"
        elif index ==3:
            self.setText()
    def setText(self):
        #text, ok = QInputDialog.getText(self, "自定义高程系统", "自定义名称:", QLineEdit.Normal)
        text, ok = QInputDialog.getText(self, "自定义高程系统",
                                        "自定义名称:", QLineEdit.Normal, self.combobox.itemText(3))
        if ok and text != '':
            self.select = text
            self.combobox.addItem(text)
            self.combobox.removeItem(3)

class Welcome_Tab(QWidget):
    def __init__(self,message):
        super(Welcome_Tab,self).__init__(parent=None)
        self.load_file_button = QPushButton(message)
        self.load_file_button.setStyleSheet("""
               QPushButton
               {
                   background-color: rgb(255, 255,255);
                   border:10px solid rgb(0, 170, 255);
               }
               """ )
        main_Layout = QFormLayout()
        main_Layout.addRow(self.load_file_button)
        self.setLayout(main_Layout)


class Current_Tab(QWidget):
    def __init__(self, parent=None):
        super(Current_Tab, self).__init__(parent)
        global c
        c = {}
        self.layout_element()
        self.errorMessageDialog = QErrorMessage(self)
        self.options = QFileDialog.Options()
        self.beding_current = []

        self.welcome_init_tab()

    def bed_group_Layout(self):
        r_group_layout = QFormLayout()
        self.pre_bed = QPushButton('导入文件', clicked=self.open_bed_file)
        self.pre_bed.setToolTip('载入潮流文件应包括时间、各层的流速数据（两种格式均可），其中流速单位为m/s')
        """self.check_multifile_bed = QCheckBox("多个文件")
        self.check_multifile_bed.setToolTip('按照所设置的高度同时转化多个文件')"""
        self.if_vd = QCheckBox("V-D格式")
        self.if_vd.setChecked(False)
        self.if_vd.setToolTip("流速为V-D格式时选择此项，为东北分向流速时取消选择")

        self.bed_count = QComboBox()
        self.bed_count.addItem('六层法')
        self.bed_count.addItem('三层法（尚未开发）')
        self.first_bin = QDoubleSpinBox()
        self.first_bin.setToolTip("输入首层高度，单位为米")
        self.bin_height = QDoubleSpinBox()
        self.bin_height.setValue(1.0)
        self.bin_height.setToolTip('设置单层高度，单位为米')

        self.top_ratio = QDoubleSpinBox()
        self.top_ratio.setValue(1.05)
        self.top_ratio.setRange(1, 1.2)
        self.top_ratio.setSingleStep(0.01)

        self.button_ratio = QDoubleSpinBox()
        self.button_ratio.setValue(0.95)
        self.button_ratio.setRange(0.9, 1)
        self.button_ratio.setSingleStep(0.01)

        self.if_save_VD = QCheckBox('保存为格式')
        self.if_save_VD.setChecked(True)
        self.save_bedded_file = QPushButton('保存结果', clicked=self.save_bedded_file_click)

        r_group_layout.addRow(self.pre_bed)
        r_group_layout.addRow(self.if_vd)
        # r_group_layout.addRow(QLabel("总层数选择"))
        # r_group_layout.addRow(self.bed_count)
        r_group_layout.addRow("首层高度", self.first_bin)
        r_group_layout.addRow("每层高度", self.bin_height)
        r_group_layout.addRow("表层流速比例", self.top_ratio)
        r_group_layout.addRow("底层流速比例", self.button_ratio)
        r_group_layout.addRow(self.save_bedded_file)
        self.bed_group_layout = r_group_layout

    def layout_element(self):

        self.current_l_Group = QGroupBox('打开潮流文件')
        self.current_m_Group = QGroupBox('文件预处理')
        self.bed_Group = QGroupBox("文件分层")

        self.current_file_button = QPushButton("导入潮流文件", clicked=self.open_file)

        self.check_multifile = QCheckBox("多个文件")
        self.check_multifile.setChecked(True)
        self.del_data = QPushButton("删除当前潮流数据", clicked=self.del_all_data)

        self.VD2EN = QPushButton('合成流速转为分向流速', clicked=self.vd_to_en)
        self.EN2VD = QPushButton('分向流速转为合成流速', clicked=self.ne_to_vd)

        l_group_layout = QVBoxLayout()
        l_group_layout.addWidget(self.current_file_button)
        l_group_layout.addWidget(self.VD2EN)
        l_group_layout.addWidget(self.EN2VD)
        l_group_layout.addWidget(self.check_multifile)
        l_group_layout.addWidget(self.del_data)
        self.current_l_Group.setLayout(l_group_layout)

        m_group_layout = QVBoxLayout()

        self.bed_group_Layout()
        self.bed_Group.setLayout(self.bed_group_layout)

        m_group_layout.addWidget(self.bed_Group)
        self.current_m_Group.setLayout(m_group_layout)

        self.sites_tab = QTabWidget()

        mainLayout = QHBoxLayout()
        mainLayout.addWidget(self.current_l_Group)
        mainLayout.addWidget(self.current_m_Group)
        mainLayout.addWidget(self.sites_tab)
        self.setLayout(mainLayout)

    def welcome_init_tab(self):
        self.init_tab = Welcome_Tab("请载入潮流数据文件")
        self.sites_tab.addTab(self.init_tab, '未载入潮流数据')

    def open_bed_file(self):
        self.pre_bed_fileName, _ = QFileDialog.getOpenFileName(self, "选取待分层文件", "E:", "CSV文档(*.csv);;Excel文件(*.xlsx)",
                                                               options=self.options)
        self.beding_current = Current(filename=self.pre_bed_fileName, is_VD=self.if_vd.isChecked())
        sig.msg_to_show.emit('文件载入成功')

    def save_bedded_file_click(self):
        fileName, _ = QFileDialog.getSaveFileName(self, '保存分层文件路径', self.pre_bed_fileName, "CSV文档(*.csv)",
                                                  options=self.options)

        self.beding_current.all_fenceng(bin=self.bin_height.value(), first_bin=self.first_bin.value(),
                                        top_ratio=self.top_ratio.value(), button_ratio=self.button_ratio.value())
        self.beding_current.save_result(outfile=fileName, V_D=self.if_save_VD)
        QMessageBox.information(self, '通知',
                                "文件转化结束")
        sig.msg_to_show.emit("文件转化结束")

    def open_file(self):
        if self.check_multifile.isChecked():
            fileName, _ = QFileDialog.getOpenFileNames(self, "选取文件", "E:",
                                                       "97-03表格文档 (*.xls);;Excel文件(*.xlsx)", options=self.options)
        else:
            fileName, _ = QFileDialog.getOpenFileName(self, "选取文件", "E:",
                                                      "97-03表格文档 (*.xls);;Excel文件(*.xlsx)", options=self.options)
            fileName = [fileName]
        return fileName

    def del_all_data(self):
        MESSAGE = "将删除所有已载入的潮流数据，是否确定？"
        reply = QMessageBox.question(self, "删除潮位数据",
                                     MESSAGE,
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            c.clear()
        else:
            pass
        sig.msg_to_show.emit("所载入潮流数据已全部删除")

    def vd_to_en(self):
        options = QFileDialog.Options()
        vd_fileName, _ = QFileDialog.getOpenFileName(self, "选取流速流向格式潮流文件", "E:", "csv文档 (*.csv);;Excel文件(*.xlsx)",
                                                     options=options)
        ne_fileName, _ = QFileDialog.getSaveFileName(self, "保存为分向流速格式潮流文件", dir_of_file(vd_fileName), "csv文档 (*.csv)",
                                                     options=options)
        vd = Current(filename=vd_fileName, is_VD=True)
        vd.convert_to_EN_file(ne_fileName)
        QMessageBox.information(self, '通知',
                                "文件转化结束")

    def ne_to_vd(self):
        options = QFileDialog.Options()
        ne_fileName, _ = QFileDialog.getOpenFileName(self, "选取分向流速格式潮流文件", "E:", "csv文档 (*.csv);;Excel文件(*.xlsx)",
                                                     options=options)
        vd_fileName, _ = QFileDialog.getSaveFileName(self, "保存为流速流向格式潮流文件", dir_of_file(ne_fileName), "csv文档 (*.csv)",
                                                     options=options)
        ne = Current(filename=ne_fileName, is_VD=False)
        ne.convert_to_VD_file(vd_fileName)
        QMessageBox.information(self, '通知',
                                "文件转化结束")


class Class_Site_Tab(QWidget):
    def __init__(self, sitename, filename, parent=None):
        super(Class_Site_Tab, self).__init__(parent)
        self.layout_element()

    def layout_element(self):
        self.lu_Group = QGroupBox("测点位置")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = M_window()

    ex.show()
    sys.exit(app.exec_())