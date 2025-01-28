import os
import xlrd
import openpyxl
from dotenv import load_dotenv
from datetime import date, datetime
from consts import requests_types
from app.jiraConn import JiraConn

load_dotenv()

#check for new files in the folder | register files in database

class Fileregister:
    '''Register DRS in database'''
    def __init__(self):
        self.filesPath = os.environ.get('DIREDITABLEFILES')
        listfiles = os.listdir(self.filesPath)
        self.files = listfiles


    def filesInDataBase(self):

        # Close file (open file - with )

        '''list of files in db - OUT(DRS Nº_VERSION)'''
        dbPath = os.environ.get('DIRDATABASE') #development
        databaseSheet = xlrd.open_workbook(dbPath).sheet_by_index(0)

        filesIndb = []
        for x in range(databaseSheet.nrows):
            drsfileN = databaseSheet.cell_value(x, 0)
            drsfileE = databaseSheet.cell_value(x, 1)
            if drsfileN != 'DRS Nº':
                NumberAndEd = f"{int(drsfileN)}_{int(drsfileE)}"
                filesIndb.append(NumberAndEd)
        reversed_list = filesIndb[::-1]
        return reversed_list


    def lastDrsCreated(self):
        print('     -> Run lastDRS...')
        allDrs = self.filesInDataBase()
        last = []
        for files in allDrs:
            drsNumber = files.split('_')[0]
            if drsNumber not in last:
                last.append(int(drsNumber))
        
        nextdrs = int(max(last)) + 1
        root = os.environ.get('ROOT')
        txtInDIr = ''
        for x in os.listdir(root):
            if '.txt' in x:
                txtInDIr = x
        if txtInDIr:
            os.remove(f'{root}/{txtInDIr}')
        j = open(f'{root}NEXT DRS_{nextdrs}.txt', 'w')
        j.close()


    def newDrsFiles(self):
        '''check for new files in folder - Returns a list of new files to add in db'''
        NewDrsInFolder = []
        # Get the full file paths and sort by modification time (descending)
        sorted_files = sorted(
            self.files,
            key=lambda f: os.path.getmtime(os.path.join(self.filesPath, f)),
            reverse=True
        )
        counter = 0
        for file in sorted_files: # number over 1000 - dont use reversed_files
            editfile = file.split('_')
            editfileName = ''.join(editfile[1].split())
            editfileVersion = ''.join(editfile[3][1].split())
            numberAndEd = f'{editfileName}_{editfileVersion}'
            if numberAndEd not in self.filesInDataBase():
                if '~$' not in file:
                    NewDrsInFolder.append(file)
                    break
            counter += 1
            if counter >=60:
                break
        return NewDrsInFolder


    def registFiles(self):
        print('     -> Run regist DRS...')
        new_drs_files = self.newDrsFiles()
        if new_drs_files:
            print('Files to Register: ', len(new_drs_files))
            for drsfiles in new_drs_files: 
                print('Register...: ', drsfiles)
                filePath = f"{os.environ.get('DIREDITABLEFILES')}{drsfiles}"

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
                    str(fileSheet.cell_value(36, 3)),
                    str(fileSheet.cell_value(37, 3)),
                    str(fileSheet.cell_value(38, 3)),
                    str(fileSheet.cell_value(39, 3)),
                    str(fileSheet.cell_value(40, 3)),
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
                drsCliente = fileSheet.cell_value(19, 3)

                #Write values in dataBase

                #Open file
                dbFile = openpyxl.load_workbook(os.environ.get('DIRDATABASE'))
                #active sheet
                dbsheet = dbFile.active
                
                emptyRowNumber = ''
                for cell in dbsheet['A']:
                    if cell.value is None:
                        emptyRowNumber = cell.row
                        break

                # add data to database
                dbsheet[f'A{emptyRowNumber}'] = drsNumber
                dbsheet[f'B{emptyRowNumber}'] = drsVersion
                dbsheet[f'C{emptyRowNumber}'] = drsCodModel
                dbsheet[f'D{emptyRowNumber}'] = drsModelName
                dbsheet[f'E{emptyRowNumber}'] = drsTipologia
                dbsheet[f'F{emptyRowNumber}'] = drsTiposPedido[0]
                dbsheet[f'G{emptyRowNumber}'] = drsTiposPedido[1]
                dbsheet[f'H{emptyRowNumber}'] = drsTiposPedido[2]


                dbsheet[f'I{emptyRowNumber}'] = drsEmpresa
                dbsheet[f'J{emptyRowNumber}'] = drsMia
                dbsheet[f'K{emptyRowNumber}'] = drsCat
                dbsheet[f'L{emptyRowNumber}'] = drsMercado
                dbsheet[f'M{emptyRowNumber}'] = drsCliente
                dbsheet[f'N{emptyRowNumber}'] = drsReqClient[0]
                dbsheet[f'O{emptyRowNumber}'] = drsReqClient[1]
                dbsheet[f'P{emptyRowNumber}'] = drsReqClient[2]
                dbsheet[f'Q{emptyRowNumber}'] = ', '.join(drsRevs)
                dbsheet[f'R{emptyRowNumber}'] = drsDeadline
                dbsheet[f'S{emptyRowNumber}'] = drsForecast
                dbsheet[f'T{emptyRowNumber}'] = drsDestino
                dbsheet[f'U{emptyRowNumber}'] = drsAprov
                dbsheet[f'V{emptyRowNumber}'] = drsAprovDate
                dbsheet[f'W{emptyRowNumber}'] = drsEspuma
                dbsheet[f'X{emptyRowNumber}'] = drsFuncRec
                dbsheet[f'Y{emptyRowNumber}'] = drsComents
                dbsheet[f'Z{emptyRowNumber}'] = drsFeedback
                dbsheet[f'AA{emptyRowNumber}'] = date.today()


                # save file and close!
                dbFile.save(os.environ.get('DIRDATABASE'))
                dbFile.close()

            now = datetime.now()
            drsTiposPedidos = ', '.join(drsTiposPedido)
            drsTiposPedidos = self.joinRequests(drsTiposPedidos)
                    

            data = {
                'drsNumber': drsNumber,
                'drsVersion':drsVersion,
                'client':drsCliente,
                'tiposPedido': drsTiposPedidos,
                'modelName': drsModelName,
                'modelCode': drsCodModel,
                'tipologia': drsTipologia,
                'deadline': drsDeadline,
            }

            # add task to jira
            jiracon = JiraConn()
            resp = jiracon.run_jira(data=data)
            if resp:
                data['issue_id'] = resp

            # send post notification to discord bot at /newtask port 8080 (localhost) with drs number / tipos pedido

            import requests

            #marge all tipos de pedido
            


            url = 'http://localhost:8080/newtask'
            
            response = requests.post(url, json=data)
            print(response.status_code)

            


            print(f'New files add to DB! at: {now}')
        else:
            return 'No new files found to add in DB!'
        
    

    #helpers
    def joinRequests(self, requests: str):
        requests = requests.split(', ')
        request_bucket = []

        for i in requests:
            if not i.strip():  # Skip empty strings
                continue
            if '(' in i:
                new_i = i.replace('(', '').replace(')', '').replace('+', ',').strip()
                request_bucket.extend([y.strip() for y in new_i.split(',')])
            else:
                request_bucket.append(i.strip())

        # Remove empty values and join with commas
        result = ','.join([req for req in request_bucket if req])
        return result.strip(',')  # Remove trailing commas


