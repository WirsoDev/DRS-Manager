#make reports by email - use AQsmtpAPI
from openpyxl import load_workbook
import datetime
import xlrd
import os
from dotenv import load_dotenv

load_dotenv()

class REPORTS:
    def __init__(self):
        #Get all data from DRS DB

        db_file = os.environ.get('DIRDATABASE')
        file = load_workbook(filename=db_file)
        self.sheet = file['Folha1']

        max = self.sheet.max_row

        self.current_mounth_drs = {}

        x = 3
        while True:
            value = self.sheet[f'AA{x}'].value
            drs_number = self.sheet[f'A{x}'].value
            if value == None:
                break
            str_value = str(value)
            recived_month = str_value.split('-')[1]
            recived_year = str_value.split('-')[0]
            # remove left 0
            if recived_month[0] == '0':
                recived_month = recived_month.replace('0', '')

            date = datetime.datetime.now()
            past_month = date.month - 1
            
            drs_name = f"{self.sheet[f'A{x}'].value}_{str(self.sheet[f'B{x}'].value)}"

            #controller if is global data or pass month
            if recived_month == str(past_month):
                if recived_year == str(date.year):
                
                    new_dic = {drs_name:{

                            'codigo_modelo': self.sheet[f'C{x}'].value,
                            'tipo_pedido_1':self.sheet[f'F{x}'].value,
                            'tipo_pedido_2':self.sheet[f'G{x}'].value,
                            'tipo_pedido_3':self.sheet[f'H{x}'].value,
                            'mercado':self.sheet[f'L{x}'].value,
                            'cliente':self.sheet[f'M{x}'].value,
                            'Uni_prod':self.sheet[f'S{x}'].value,
                            'Resp_pedido':self.sheet[f'U{x}'].value,
                            'dep_requerente':self.sheet[f'V{x}'].value,
                            'finalizado_por':self.sheet[f'AB{x}'].value,
                            'aprovado_por':self.sheet[f'AD{x}'].value,
                            'recusado_por':self.sheet[f'AE{x}'].value,
                            
                            }}
                    self.current_mounth_drs.update(new_dic)

            x += 1

        file.close()

   
    def getDrsTotalNum(self):
        return len(self.current_mounth_drs)
    

    def getAllDrsData(self):
        return self.current_mounth_drs


    def getAllDrsDataList(self):
        data = self.current_mounth_drs
        for x in data:
            print(x)
            print(self.current_mounth_drs[x])
            print('-'* 20)

    
    def GetNonAproved(self):
        data = self.current_mounth_drs
        drs_bucket = []
        for x in data:
            if self.current_mounth_drs[x]['aprovado_por'] == None:
                drs_bucket.append(self.current_mounth_drs[x])
        print(drs_bucket)
        return drs_bucket
    

    def GetAproved(self):
        data = self.current_mounth_drs
        drs_bucket = []
        for x in data:
            if self.current_mounth_drs[x]['aprovado_por'] != None:
                drs_bucket.append(self.current_mounth_drs[x])
        print(drs_bucket)
        return drs_bucket

    def getMarkets(self):
        ''' Get all markets 
            return: DIC {tipo:qnt} '''

        data = self.current_mounth_drs

        tipos = []
        tipos_dic = {}
        for key, values in data.items():
            tipo = values['mercado']
            if tipo:
                tipos.append(tipo)
            
        for tip in tipos:
            y = {
                tip:0
            }
            tipos_dic.update(y)
        
        for x in tipos:
            tipos_dic[x] += 1


        dic_sorted = sorted(tipos_dic.items(), key=lambda x: x[1], reverse=True)
        print(dic_sorted)
        return dic_sorted 
    

    def getClients(self):
        ''' Get all markets 
            return: DIC {tipo:qnt} '''

        data = self.current_mounth_drs

        tipos = []
        tipos_dic = {}
        for key, values in data.items():
            tipo = values['cliente']
            if tipo:
                tipos.append(tipo)
            
        for tip in tipos:
            y = {
                tip:0
            }
            tipos_dic.update(y)
        
        for x in tipos:
            tipos_dic[x] += 1


        dic_sorted = sorted(tipos_dic.items(), key=lambda x: x[1], reverse=True)
        print(dic_sorted)
        return dic_sorted
    

    def getTypeRequest(self):
        ''' Get all markets 
            return: DIC {tipo:qnt} '''

        data = self.current_mounth_drs

        tipos = []
        tipos_dic = {}
        for key, values in data.items():
            tipo = values['tipo_pedido_1']
            if tipo != None:
                tipos.append(tipo)
            tipo_2 = values['tipo_pedido_2']
            if tipo_2 != None:
                tipos.append(tipo_2)
            tipo_3 = values['tipo_pedido_3']
            if tipo_3 != None:
                tipos.append(tipo_3)
            
        for tip in tipos:
            y = {
                tip:0
            }
            tipos_dic.update(y)
        
        for x in tipos:
            tipos_dic[x] += 1


        dic_sorted = sorted(tipos_dic.items(), key=lambda x: x[1], reverse=True)
        print(dic_sorted)
        return dic_sorted 
    



    


    


x = REPORTS()
x.getTypeRequest()
