import os
import xlrd
import openpyxl
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
        if self.newDrsFiles():
            print('Files to Register: ', len(self.newDrsFiles()))
            print(self.newDrsFiles())
            for drsfiles in self.newDrsFiles(): #test!!!!! ----------------- remove de comments!
                filePath = f"{os.environ.get('DIREDITABLEFILES')}\{drsfiles}"

                #Extract values from files
                fileSheet = xlrd.open_workbook(filePath).sheet_by_index(0)
                drsNumber = int(fileSheet.cell_value(3, 0))
                drsVersion = int(fileSheet.cell_value(3, 3))
                drsCodModel = fileSheet.cell_value(6, 0)
                drsModelName = fileSheet.cell_value(6, 3)
                drsTipologia = fileSheet.cell_value(6, 6)
                drsTiposPedido = [
                    fileSheet.cell_value(8, 3),
                    fileSheet.cell_value(8, 5),
                    fileSheet.cell_value(8, 7)
                    ]
                drsEmpresa = fileSheet.cell_value(11, 3)
                drsMia = fileSheet.cell_value(13, 3)
                drsCat = fileSheet.cell_value(15, 3)
                drsMercado = fileSheet.cell_value(17, 3)
                drsCliente = fileSheet.cell_value(19, 4)
                drsReqClient = [
                    fileSheet.cell_value(22, 3),
                    fileSheet.cell_value(24, 3),
                    fileSheet.cell_value(27, 0)
                    ]
                drsRevs = [
                    [fileSheet.cell_value(36, 0), fileSheet.cell_value(36, 3)],
                    [fileSheet.cell_value(37, 0), fileSheet.cell_value(37, 3)],
                    [fileSheet.cell_value(38, 0), fileSheet.cell_value(38, 3)],
                    [fileSheet.cell_value(39, 0), fileSheet.cell_value(39, 3)],
                    [fileSheet.cell_value(40, 0), fileSheet.cell_value(40, 3)],
                ]
                drsDeadline = fileSheet.cell_value(11, 7)
                drsForecast = fileSheet.cell_value(13, 7)
                drsDestino = fileSheet.cell_value(15, 7)
                drsAprov = fileSheet.cell_value(17, 7)
                drsAprovDate = fileSheet.cell_value(19, 7)
                drsEspuma = fileSheet.cell_value(22, 5)
                drsFuncRec = fileSheet.cell_value(25, 5)
                drsComents = fileSheet.cell_value(28, 5)
                drsFeedback = fileSheet.cell_value(33, 5)

                #Write values in dataBase

                #Open file
                dbFile = openpyxl.load_workbook(os.environ.get('DIRDATABASE'))
                #active sheet
                dbsheet = dbFile.active

                # To write values
                #dbsheet['C3'] = 3
                #dbFile.save(os.environ.get('DIRDATABASE'))

                #check whats the prox enpty row
                emptyRowNumber = ''
                for cell in dbsheet['A']:
                    if cell.value is None:
                        emptyRowNumber = cell.row
                        break
                # add data to database

                dbsheet[f'A{emptyRowNumber}'] = drsNumber
                dbsheet[f'B{emptyRowNumber}'] = drsVersion
                dbFile.save(os.environ.get('DIRDATABASE'))
                dbFile.close()
                print('New files add to DB!')

        else:
            return 'No new files found!'


