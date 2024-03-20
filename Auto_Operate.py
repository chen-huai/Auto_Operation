import sys
import os
import re
import time
import math
import pandas as pd
import csv
import numpy as np
import win32com.client
import datetime
import chicon  # 引用图标
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QMessageBox, QFileDialog
# from PyQt5.QtCore import *
from Get_Data import *
from File_Operate import *
# from PDF_Operate import *
from Auto_Operate_Ui import Ui_MainWindow
from Data_Table import *
from Logger import *
from Windows_Operate import *


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionExit.triggered.connect(MyMainWindow.close)
        self.actionHelp.triggered.connect(self.showVersion)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.pushButton_12.clicked.connect(self.textBrowser.clear)
        self.pushButton_16.clicked.connect(self.getFileUrl)
        self.pushButton_49.clicked.connect(self.viewData)
        self.pushButton_56.clicked.connect(self.getCoordinates)
        self.pushButton_57.clicked.connect(self.win_auto_step)
        self.filesUrl = []

    def getConfig(self):
        # 初始化，获取或生成配置文件
        global configFileUrl
        global desktopUrl
        global now
        global last_time
        global today
        global oneWeekday
        global fileUrl

        date = datetime.datetime.now() + datetime.timedelta(days=1)
        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y.%m.%d')
        oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s/config' % desktopUrl
        configFile = os.path.exists('%s/config_auto.csv' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                if not os.path.exists(configFileUrl):
                    os.makedirs(configFileUrl)
                MyMainWindow.createConfigContent(self)
                MyMainWindow.getConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
            else:
                exit()
        else:
            MyMainWindow.getConfigContent(self)

    # 获取配置文件内容
    def getConfigContent(self):
        # 配置文件
        csvFile = pd.read_csv('%s/config_auto.csv' % configFileUrl, names=['A', 'B', 'C'])
        global configContent
        global username
        global role
        configContent = {}
        username = list(csvFile['A'])
        number = list(csvFile['B'])
        role = list(csvFile['C'])
        for i in range(len(username)):
            configContent['%s' % username[i]] = number[i]

        try:
            self.textBrowser.append("配置获取成功")
        except AttributeError:
            QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
        else:
            pass

    # 创建配置文件
    def createConfigContent(self):
        global monthAbbrev
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        configContent = [
            ['特殊开票', '内容', '备注'],
            ['Step_Data_URL', desktopUrl, '流程步骤数据路径'],
            ['文件说明', '序号，步骤，备注', '文件列头说明'],
            ['步骤说明', 'move,x,y,time', '移动鼠标'],
            ['步骤说明', 'click,x,y', '点击鼠标'],
            ['步骤说明', 'dragTo,x,y,time', '拖动鼠标'],
            ['步骤说明', 'mouseDown,x,y,button', '按下鼠标'],
            ['步骤说明', 'mouseUp,x,y,button', '释放鼠标'],
            ['步骤说明', 'typewrite,****', '键盘输入'],
            ['步骤说明', 'press,button', '模拟按下并释放按键'],
            ['步骤说明', 'hotkey,button1,button2', '模拟按下热键并释放按键'],
            ['步骤说明', 'sleep,time', '暂停时间'],
        ]
        config = np.array(configContent)
        df = pd.DataFrame(config)
        df.to_csv('%s/config_auto.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_auto.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
                                QMessageBox.Yes)

    # 导出配置文件
    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    # 导入配置文件
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.getConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

    def showAuthorMessage(self):
        # 关于作者
        QMessageBox.about(self, "关于",
                          "人生苦短，码上行乐。\n\n\n        ----Frank Chen")

    def showVersion(self):
        # 关于作者
        QMessageBox.about(self, "版本",
                          "V 24.01\n\n\n 2024-02-01")

    # 获取坐标
    def getCoordinates(self):
        win_obj = Windows_auto()
        time = int(self.spinBox.text())
        coordinates = win_obj.get_desktop_coordinates(time)
        self.textBrowser.append("坐标：(%s,%s)" % (coordinates[0], coordinates[1]))
        return coordinates

    # 获取文件
    def getFile(self):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件',
                                                      '%s\\%s' % (configContent['Step_Data_URL'], today),
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl

    # 步骤数据路径
    def getFileUrl(self):
        fileUrl = MyMainWindow.getFile(self)
        if fileUrl:
            self.lineEdit_6.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    # 查看SAP操作数据详情
    def viewData(self):
        try:
            fileUrl = self.lineEdit_6.text()
            odm_data_obj = Get_Data()
            df = odm_data_obj.getFileData(fileUrl)
            myTable.createTable(df)
            myTable.showMaximized()
        except Exception as errorMsg:
            self.textBrowser.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            app.processEvents()
            return

    # 获取b步骤数据
    def getStepListData(self):
        try:
            Step_Data_URL = self.lineEdit_6.text()
            if Step_Data_URL == '':
                self.textBrowser.append('无选中文件')
                self.textBrowser.append('----------------------------------')
                app.processEvents()
                return None
            else:
                step_list_obj = Get_Data()
                step_list_data = step_list_obj.getFileData(Step_Data_URL)
                return step_list_data
        except Exception as errorMsg:
            self.textBrowser.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            app.processEvents()
            return

    # 电脑桌面操作
    def win_auto_step(self):
        steps_list_data = myWin.getStepListData()
        # steps_num = list(steps_list_data['序号'])
        steps = list(steps_list_data['步骤'])
        win_auto_obj = Windows_auto()
        cycles_num = int(self.spinBox_2.text())
        cycle_msg = {}
        step_msg = {}
        for cycle in range(cycles_num):
            cycle_msg['flag'] = 1
            self.textBrowser.append('第%s次循环步骤' % (cycle+1))
            flag = 1
            if self.checkBox.isChecked():
                reply = QMessageBox.question(self, '信息', '是否继续操作电脑桌面！！！',
                                             QMessageBox.Yes | QMessageBox.No,
                                             QMessageBox.Yes)
                if reply == QMessageBox.No:
                    flag = 0
            if flag == 1:
                num = 1
                step_msg['flag'] = 1
                for step in steps:
                    step_flag = 1
                    step_list = step.split(',')
                    step_flag = win_auto_obj.window_automation_operation(step_list)
                    if step_flag == 0:
                        step_msg['flag'] = 0
                        cycle_msg['flag'] = 0
                        step_msg['information'] = '第%s步骤失败' % num
                        self.textBrowser.append("<font color='red'> 第%s步骤失败 </font>" % num)
                        app.processEvents()
                    num += 1
            else:
                self.textBrowser.append('跳过第%s次循环步骤' % (cycle + 1))
            if cycle_msg['flag'] == 1:
                self.textBrowser.append('成功')
                app.processEvents()
            else:
                self.textBrowser.append("<font color='red'> 失败 </font>" % num)



if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())
