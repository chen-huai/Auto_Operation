# from PyQt5 import QtCore, QtGui, QtWidgets
import os.path

from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QItemDelegate
from Table_Ui import *
import pandas as pd


class MyTableWindow(QMainWindow, Ui_TableWindow):
    def __init__(self, parent=None):
        super(MyTableWindow, self).__init__(parent)
        self.setupUi(self)
        # self.pushButton.clicked.connect(self.saveTable)
        # self.pushButton_2.clicked.connect(self.createTable)

    def createTable(self, df):
        self.df = df
        self.df = self.df.astype(str)
        self.df_rows = self.df.shape[0]
        self.df_cols = self.df.shape[1]
        self.tableWidget.setRowCount(self.df_rows)
        self.tableWidget.setColumnCount(self.df_cols)

        ##设置水平表头
        self.tableWidget.setHorizontalHeaderLabels(self.df.keys().astype(str))

        # self.tabletWidget.
        for i in range(self.df_rows):
            for j in range(self.df_cols):
                self.tableWidget.setItem(i, j, QTableWidgetItem(self.df.iloc[i, j]))
        # 第1列不允许编辑
        self.tableWidget.setItemDelegateForColumn(0, EmptyDelegate(self))
        # 行颜色
        self.tableWidget.setAlternatingRowColors(True)
        # 显示所有内容
        self.tableWidget.resizeColumnsToContents()
        # 平均分配
        self.tableWidget.horizontalHeader().setSectionResizeMode(True)
        # 排序
        self.tableWidget.setSortingEnabled(True)
        # 设置默认第几列排序，0列开始
        # self.tableWidget.sortItems(0)

    @pyqtSlot()
    def print_my_df(self):
        print(self.df)

    # @pyqtSlot()
    # def saveTable(self, filePath):
    #     col = self.tableWidget.columnCount()
    #     row = self.tableWidget.rowCount()
    #     # for currentQTableWidgetItem in self.tableWidget.selectedItems():
    #     # 	print((currentQTableWidgetItem.row(), currentQTableWidgetItem.column()))
    #     data = []
    #     for i in range(col):
    #         data.append(i)
    #         data[i] = []
    #         for j in range(row):
    #             itemData = self.tableWidget.item(j, i).text()
    #             data[i].append(itemData)
    #     configFile = pd.DataFrame({'a': data[0], 'b': data[1], 'c': data[2]})
    #     configFile.to_excel('%s' % filePath, encoding="utf_8_sig", index=0, header=0)
    #     reply = QMessageBox.question(self, '信息', '配置文件已修改成功，是否重新获取新的config文件内容',
    #                                  QMessageBox.Yes | QMessageBox.No,
    #                                  QMessageBox.Yes)
    #     if reply == QMessageBox.Yes:
    #         MyMainWindow.getConfigContent(self)


# table不可编辑
class EmptyDelegate(QItemDelegate):
    def __init__(self, parent):
        super(EmptyDelegate, self).__init__(parent)

    def createEditor(self, QWidget, QStyleOptionViewItem, QModelIndex):
        return None
