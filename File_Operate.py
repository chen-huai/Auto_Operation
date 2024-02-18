import time
import os

class File_Opetate():
    def createFolder(self, url):
        isExists = os.path.exists(url)
        if not isExists:
            os.makedirs(url)
    def getFileName(self, fileUrl, fileName, fileType):
        nowTime = time.strftime('%Y-%m-%d %H.%M.%S')
        fileName = fileUrl + '/' + nowTime + ' - ' + fileName + '.' + fileType
        return fileName

