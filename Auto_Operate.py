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
        self.pushButton_49.clicked.connect(self.viewOdmData)
        self.pushButton_56.clicked.connect(lambda: self.orderUnlockOrLock('Unlock'))
        self.pushButton_57.clicked.connect(lambda: self.orderUnlockOrLock('Lock'))
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
        configFile = os.path.exists('%s/config_sap_controlling.csv' % configFileUrl)
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
        csvFile = pd.read_csv('%s/config_sap_controlling.csv' % configFileUrl, names=['A', 'B', 'C'])
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
            ['SAP_Date_URL', 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM',
             '文件数据路径'],
        ]
        config = np.array(configContent)
        df = pd.DataFrame(config)
        df.to_csv('%s/config_sap_controlling.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_sap_controlling.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
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


    # 获取文件
    def getFile(self):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件',
                                                      '%s\\%s' % (configContent['SAP_Date_URL'], today),
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl

    # SAP数据路径
    def getFileUrl(self):
        fileUrl = MyMainWindow.getFile(self)
        if fileUrl:
            self.lineEdit_6.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)


    # 获取billing list文件
    def getBillingListFile(self):
        try:
            selectBatchFile = QFileDialog.getOpenFileName(self, '选择文件',
                                                          '%s' % configContent['Billing_List_URL'],
                                                          'files(*.xlsx)')
            fileUrl = selectBatchFile[0]
            if fileUrl:
                self.lineEdit_25.setText(fileUrl)
                app.processEvents()
                self.textBrowser_3.append('选中Billing List文件：%s' % fileUrl)
                self.textBrowser_3.append('----------------------------------')
            else:
                self.textBrowser_3.append('无选中文件')
                self.textBrowser_3.append('----------------------------------')
            app.processEvents()
            return fileUrl
        except Exception as errorMsg:
            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            app.processEvents()
            return

    # 获取billing数据
    def getBillingListData(self, sheet_name=[]):
        try:
            billing_list_url = self.lineEdit_25.text()
            if billing_list_url == '':
                self.textBrowser_3.append('无选中文件')
                self.textBrowser_3.append('----------------------------------')
                app.processEvents()
                return None
            else:
                billing_list_obj = Get_Data()
                billing_list_data = billing_list_obj.getFileMoreSheetData(billing_list_url, sheet_name)
                pivotTableKey = ['Final Invoice No.', 'Customer Name', 'CS', 'Cur.']
                valusKey = ['求和项:Amount with VAT']
                billing_list_data = billing_list_obj.pivotTable(pivotTableKey, valusKey)
                billing_list_data = billing_list_data.reset_index()
                return billing_list_data
        except Exception as errorMsg:
            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            app.processEvents()
            return


    # 查看Billing list数据
    def viewBillingListData(self):
        fileUrl = self.lineEdit_25.text()
        if fileUrl:
            sheet_name = []
            df = myWin.getBillingListData(sheet_name)
            try:
                myTable.createTable(df)
                myTable.showMaximized()
            except Exception as errorMsg:
                self.textBrowser_3.append('数据有问题%s' % errorMsg)
                self.textBrowser_3.append('----------------------------------')
                app.processEvents()
        else:
            self.textBrowser_3.append('无选中文件')
            self.textBrowser_3.append('----------------------------------')
            app.processEvents()


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())
