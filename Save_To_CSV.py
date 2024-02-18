import pandas as pd
import numpy as np
import os

class Save_To_CSV():
    def __init__(self):

        pass

    def createFolder(self, url):
        isExists = os.path.exists(url)
        if not isExists:
            os.makedirs(url)

    def fileName(self, file):
        (filepath, filename) = os.path.split(file)
        return filepath,filename

    def saveToCsv(self, data, fileUrl, filename):
        csvData = pd.DataFrame(data)
        csvData.to_csv('%s\\%s' % (fileUrl, filename), mode='a', index=0, encoding='utf_8_sig')






path = 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data\\ODM Raw Data\\..\\Raw Data\\20220411'
newCsv = Save_To_CSV()
newCsv.createFolder(path)
data = {
    '索引': [],
    'Project No.': [],
    'Sales': [],
    'Material Code': [],
    "Invoices' name (Chinese)": [],
    'Currency': [],
    'Exchange Rate': [],
    'Month': [],
    'CS': [],
    'Buyer(GPC)': [],
    'Amount': [],
    'Amount with VAT': [],
    'Total Cost': [],
    'num': [],
    'Text': [],
    'Long Text': [],
    'Order No.': [],
    'Proforma No.': [],
    'Update Time': [],
}
data['索引'].append(1)
data['Project No.'].append(12)
data['Sales']. append(13)
data['Material Code']. append(14)
data["Invoices' name (Chinese)"]. append(15)
data['Currency']. append(16)
data['Exchange Rate']. append(17)
data['Month']. append(18)
data['CS']. append(19)
data['Buyer(GPC)']. append(111)
data['Amount']. append(112)
data['Amount with VAT']. append(113)
data['Total Cost']. append(114)
data['num']. append(115)
data['Text']. append(116)
data['Long Text']. append(117)
data['Order No.']. append(118)
data['Proforma No.']. append(119)

newCsv.saveToCsv(data, path, 'test.csv')