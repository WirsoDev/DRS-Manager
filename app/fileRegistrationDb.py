import os
import xlrd
from dotenv import load_dotenv

load_dotenv()

#check for new files in the folder | register files in database

class Fileregister:
    '''Register DRS in database'''
    def __init__(self):
        filesPath = os.environ.get('DIREDITABLEFILES')
        listfiles = os.listdir(filesPath)
        self.files = listfiles


    def filesInDataBase(self):
        '''list of files in db - OUT(DRS Nº_VERSION)'''
        dbPath = os.environ.get('DIRDATABASE')
        databaseSheet = xlrd.open_workbook(dbPath).sheet_by_index(0)

        filesIndb = []
        for x in range(databaseSheet.nrows):
            drsfileN = databaseSheet.cell_value(x, 0)
            drsfileE = databaseSheet.cell_value(x, 1)
            if drsfileN != 'DRS Nº':
                NumberAndEd = f"{int(drsfileN)}_{int(drsfileE)}"
                filesIndb.append(NumberAndEd)
        return filesIndb


    def newDrsFiles(self):
        '''check for new files in folder - Returns a list of new files to add in db'''
        NewDrsInFolder = []
        for file in self.files:
            editfile = file.split('_')
            editfileName = ''.join(editfile[1].split())
            editfileVersion = ''.join(editfile[3][1].split())
            numberAndEd = f'{editfileName}_{editfileVersion}'
            if numberAndEd not in self.filesInDataBase():
                NewDrsInFolder.append(file)
        return NewDrsInFolder


    def registFiles(self):
        pass

    
    
