import pandas as pd
import numpy as np

class Get_Data():
    # def __init__(self,fileDataUrl):
    #     self.fileDataUrl = fileDataUrl
        # self.getFileData()
        # self.getHeaderData()
        # self.getIndexNumForHead()
        # self.getFileDataList()

    def getFileData(self, fileDataUrl):
        self.fileDataUrl = fileDataUrl
        fileType = self.fileDataUrl.split(".")[-1]
        if fileType == 'xlsx':
            self.fileData = pd.read_excel(self.fileDataUrl)
            # self.fileData = pd.read_excel(self.fileDataUrl, dtype='str')
            # self.fileData = pd.read_excel(self.fileDataUrl, keep_default_na=False)
        elif fileType == 'csv':
            self.fileData = pd.read_csv(self.fileDataUrl)
            # self.fileData = pd.read_csv(self.fileDataUrl, dtype='str')
            # self.fileData = pd.read_csv(self.fileDataUrl, keep_default_na=False)
        height, width = self.fileData.shape
        return self.fileData

    def getFileMoreSheetData(self, fileDataUrl, sheet_name=[]):
        if sheet_name==[]:
            sheet_name = None
        self.fileDataUrl = fileDataUrl
        self.fileData = pd.read_excel(self.fileDataUrl, sheet_name=sheet_name)
        self.fileData = pd.concat(self.fileData.values(), ignore_index=True)
        self.fileData.dropna(subset=['Final Invoice No.'], inplace=True)
        return self.fileData

    def getMergeFileData(self, fileDataUrl):
        self.fileDataUrl = fileDataUrl
        fileType = self.fileDataUrl.split(".")[-1]
        if fileType == 'xlsx':
            # self.fileData = pd.read_excel(self.fileDataUrl)
            self.fileData = pd.read_excel(self.fileDataUrl, float_precision='round_trip', dtype='str')
            # self.fileData = pd.read_excel(self.fileDataUrl, keep_default_na=False)
        elif fileType == 'csv':
            # self.fileData = pd.read_csv(self.fileDataUrl)
            self.fileData = pd.read_csv(self.fileDataUrl, dtype='str')
            # self.fileData = pd.read_csv(self.fileDataUrl, keep_default_na=False)
        height, width = self.fileData.shape
        return self.fileData
    def getHeaderData(self):
        self.headData = list(self.fileData.head())
        return self.headData
    def getIndexNumForHead(self):
        self.projectNo = self.headData.index('Project No.')
        self.cs = self.headData.index('CS')
        self.sales = self.headData.index('Sales')
        self.currency = self.headData.index('Currency')
        self.partnerCode = self.headData.index('GPC Glo. Par. Code')
        self.materialCode = self.headData.index('Material Code')
        self.phyMaterialCode = self.headData.index('PHY Material Code')
        self.chmMaterialCode = self.headData.index('CHM Material Code')
        self.sapNo = self.headData.index('SAP No.')
        self.amount = self.headData.index('Amount')
        self.amountWithVAT = self.headData.index('Amount with VAT')
        self.exchangeRate = self.headData.index('Exchange Rate')
        self.costList = list(self.fileData['Total Cost'])
        return self.projectNo, self.cs, self.sales, self.currency, self.partnerCode, self.materialCode, self.phyMaterialCode, self.chmMaterialCode, self.sapNo, self.amount, self.amountWithVAT, self.exchangeRate,self.costList
    def deleteTheRows(self, deleteRowList = {}):
        for key in deleteRowList:
            self.fileData = self.fileData[self.fileData[key] != deleteRowList[key]]
        return self.fileData
    def fillNanColumn(self,fillNanColumnKey):
        for filledKey in fillNanColumnKey:
            for fillKey in fillNanColumnKey[filledKey]:
                self.fileData[filledKey].fillna(self.fileData[fillKey], inplace=True)
        # self.fileData["Material Code"].fillna(self.fileData["PHY Material Code"], inplace=True)
        # self.fileData["Material Code"].fillna(self.fileData["CHM Material Code"], inplace=True)
        return self.fileData
    # def pivotTable(self):
    def pivotTable(self,pivotTableKey, valusKey):
        pivotData = pd.pivot_table(self.fileData, index=pivotTableKey, values=valusKey, aggfunc='sum')
        return pivotData
    def getFileDataList(self,getFileDataListKey):
        self.fileDataList = {}
        for each in getFileDataListKey:
            self.fileDataList[each] = list(self.fileData[each])
        return self.fileDataList

    def getFileDataList1(self):
        self.fileData = self.fileData[self.fileData['Amount'] != 0]
        # self.fileData.dropna(axis=0,subset=["Amount"],inplace = True)
        self.projectNoList = list(self.fileData['Project No.'])
        self.csList = list(self.fileData['CS'])
        self.salesList = list(self.fileData['Sales'])
        self.currencyList = list(self.fileData['Currency'])
        self.partnerCodeList = list(self.fileData['GPC Glo. Par. Code'])
        self.materialCodeList = list(self.fileData['Material Code'])
        self.sapNoList = list(self.fileData['SAP No.'])
        self.amountList = list(self.fileData['Amount'])
        self.amountWithVATList = list(self.fileData['Amount with VAT'])
        self.exchangeRateList = list(self.fileData['Exchange Rate'])
        self.costList = list(self.fileData['Total Cost'])
        return self.projectNoList, self.csList, self.salesList, self.currencyList, self.partnerCodeList, self.materialCodeList,self.sapNoList, self.amountList, self.amountWithVATList, self.exchangeRateList,self.costList

    def deleteTheColumn(self, deleteColumnList):
        self.fileData.drop(labels=deleteColumnList, axis=1, inplace=True)
        return self.fileData

    def mergeData(self, data1, data2, onData):
        mergeData = pd.merge(data1, data2, on=onData, how='inner')
        return mergeData

    def column_concat_func(self, data):
        # 行信息合并
        return pd.Series({
                'combine_column_msg': '\n'.join(data['column_msg'].unique()),
            })

    def row_concat_func(self, data):
        # 列信息合并
        return pd.Series({
                'combine_row_msg': '\n'.join(data['row_msg'].unique()),
            }
            )


# data = {
#     'col1': [1, 2, 3],
#     'col2': [4, 5, 6],
#     'col3': [7, 8, 9],
#     'col4': [10, 11, 12]
# }
#
# df = pd.DataFrame(data)
#
# # 选择要合并的三列，并使用apply函数将它们相加并用制表符隔开
# # df['merged'] = df[['col1', 'col2', 'col3']].apply(lambda row: '\t'.join(map(str, row)), axis=1)
#
# df['merged'] = df[['col1', 'col3', 'col2']].apply(lambda row: '\n'.join(f"{col}:{val}" for col, val in zip(df[['col1', 'col3', 'col2']], row)), axis=1)
#
# # 移除原始三列
# df.drop(['col1', 'col2', 'col3'], axis=1, inplace=True)


