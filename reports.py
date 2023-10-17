#make reports by email - use AQsmtpAPI
from openpyxl import load_workbook
import datetime
import xlrd


class REPORTS:
    def __init__(self):
        #Get all data from DRS DB

        file = load_workbook(filename=r'./DRS_DB.xlsx')
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
                            

                            }}
                    self.current_mounth_drs.update(new_dic)

            x += 1

        file.close()


        
    def getDrsTotalNum(self):
        return len(self.current_mounth_drs)
    


    


x = REPORTS()
print(x.getDrsTotalNum())
