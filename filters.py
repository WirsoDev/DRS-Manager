from reports import REPORTS
import xlrd
import openpyxl
from datetime import datetime, date
import time
import os

class FILTER_DATA:
    #filter data from reports

    def __init__(self, global_= False, dep=None, request=None, _open=False, filter_year=False):
        self.isGobal = global_
        self.filter_year = filter_year
        self.reports = REPORTS(self.isGobal, filter_year=self.filter_year)
        self.data = self.reports.getAllDrsData()
        self.dep = dep
        self.request = request
        self.isopen = _open


    def requests_filter(self):
        requests = {
            'MRK':[
                'APRESENTAÇÕES',
                'DESENV. PROD.',
                'IMAGEM DE PRODUTO',
                'CATALOGOS/DEV.GRAFICO',
                '3D-AMBIENTE',
                'MULTI-MEDIA',
                'BOOK',
                '3D- MODELAÇÃO',
                '3D-RENDER',
                'VIDEO',
                'SHOOTING',
                'DESCRITIVO DE MARKETING'
            ],
            'CUSTEIO':[
                'CUSTEIO',
                'REVISÃO DE PREÇO'
            ],
            'DID':[
                'CERTIFICAÇÃO',
                'PRODUÇÃO DE AMOSTRA',
                'PROTOTIPO',
                'PROTOTIPO + CERTIFICAÇÃO'
                'CERTIFICAÇAO',
                'REENGENHARIA'
            ]
        }

        if self.request:
            return [self.request]
        
        if self.dep:
            return requests[self.dep]
        
        return [value for sublist in requests.values() for value in sublist]



    def get_filtered_data(self):

        data = self.data
        accepted_requests = self.requests_filter()
        DRS_FILTER = []


        for drs in data:
            if self.isopen == True:
                if data[drs]['finalizado_por'] == None:
                    drs_n = data[drs]['drs_n']
                    data_registo = data[drs]['data_registo']
                    if type(data_registo) == str:
                        new_date = datetime.strptime(data_registo, '%d/%m/%Y')
                    else:
                        new_date = data_registo
                    cliente = data[drs]['cliente']
                    tipo_pedido_1 = data[drs]['tipo_pedido_1']
                    tipo_pedido_2 = data[drs]['tipo_pedido_2']
                    tipo_pedido_3 = data[drs]['tipo_pedido_3']

                    if tipo_pedido_1 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'projeto':tipo_pedido_1,
                            'cliente':cliente,
                            'data':new_date
                        }
                        DRS_FILTER.append(new_data)
                    
                    if tipo_pedido_2 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'projeto':tipo_pedido_2,
                            'cliente':cliente,
                            'data':new_date
                        }
                        DRS_FILTER.append(new_data)
                    
                    if tipo_pedido_3 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'projeto':tipo_pedido_3,
                            'cliente':cliente,
                            'data':new_date
                        }
                        DRS_FILTER.append(new_data)  

            if self.isopen == False:
                if data[drs]['finalizado_por'] != None:
                    drs_n = data[drs]['drs_n']
                    data_registo = data[drs]['finalizado_em']
                    if type(data_registo) == str:
                        new_date = datetime.strptime(data_registo, '%d/%m/%Y')
                    else:
                        new_date = data_registo
                    cliente = data[drs]['cliente']
                    tipo_pedido_1 = data[drs]['tipo_pedido_1']
                    tipo_pedido_2 = data[drs]['tipo_pedido_2']
                    tipo_pedido_3 = data[drs]['tipo_pedido_3']

                    if tipo_pedido_1 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'projeto':tipo_pedido_1,
                            'cliente':cliente,
                            'data':new_date
                        }
                        DRS_FILTER.append(new_data)
                    
                    if tipo_pedido_2 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'projeto':tipo_pedido_2,
                            'cliente':cliente,
                            'data':new_date
                        }
                        DRS_FILTER.append(new_data)
                    
                    if tipo_pedido_3 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'projeto':tipo_pedido_3,
                            'cliente':cliente,
                            'data':new_date
                        }
                        DRS_FILTER.append(new_data)
            

            if self.isopen == 'All':
                if True:
                    drs_n = data[drs]['drs_n']

                    data_registo = data[drs]['data_registo']
                    data_finalizado = data[drs]['finalizado_em']

                    if type(data_registo) == str:
                        new_date = datetime.strptime(data_registo, '%d/%m/%Y')
                    else:
                        new_date = data_registo

                    if type(data_finalizado) == str:
                        new_date_finalizado = datetime.strptime(data_finalizado, '%d/%m/%Y')
                    else:
                        new_date_finalizado = data_finalizado
                    
                    codigo_modelo = data[drs]['codigo_modelo']
                    mercado = data[drs]['mercado']
                    Uni_prod = data[drs]['Uni_prod']
                    Resp_pedido = data[drs]['Resp_pedido']
                    dep_requerente = data[drs]['dep_requerente']
                    finalizado_por = data[drs]['finalizado_por']
                    recusado_por = data[drs]['recusado_por']
                    anulado_por = data[drs]['anulado_por']
                    aprovado_por = data[drs]['aprovado_por']
                    cliente = data[drs]['cliente']

                    tipo_pedido_1 = data[drs]['tipo_pedido_1']
                    tipo_pedido_2 = data[drs]['tipo_pedido_2']
                    tipo_pedido_3 = data[drs]['tipo_pedido_3']

                    if tipo_pedido_1 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'codigo_modelo':codigo_modelo,
                            'projeto':tipo_pedido_1,
                            'mercado':mercado,
                            'cliente':cliente,
                            'uni_prod':Uni_prod,
                            'resp_pedido':Resp_pedido,
                            'dep_requerente':dep_requerente,
                            'data_registo':new_date,
                            'finalizado_por':finalizado_por,
                            'finalizado_em':new_date_finalizado,
                            'aprovado_por':aprovado_por,
                            'recusado_por':recusado_por,
                            'anulado_por':anulado_por
                        }
                        DRS_FILTER.append(new_data)
                    
                    if tipo_pedido_2 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'codigo_modelo':codigo_modelo,
                            'projeto':tipo_pedido_2,
                            'mercado':mercado,
                            'cliente':cliente,
                            'uni_prod':Uni_prod,
                            'resp_pedido':Resp_pedido,
                            'dep_requerente':dep_requerente,
                            'data_registo':new_date,
                            'finalizado_por':finalizado_por,
                            'finalizado_em':new_date_finalizado,
                            'aprovado_por':aprovado_por,
                            'recusado_por':recusado_por,
                            'anulado_por':anulado_por
                        }
                        DRS_FILTER.append(new_data)
                    
                    if tipo_pedido_3 in accepted_requests:
                        new_data = {
                            'drs':drs_n,
                            'codigo_modelo':codigo_modelo,
                            'projeto':tipo_pedido_3,
                            'mercado':mercado,
                            'cliente':cliente,
                            'uni_prod':Uni_prod,
                            'resp_pedido':Resp_pedido,
                            'dep_requerente':dep_requerente,
                            'data_registo':new_date,
                            'finalizado_por':finalizado_por,
                            'finalizado_em':new_date_finalizado,
                            'aprovado_por':aprovado_por,
                            'recusado_por':recusado_por,
                            'anulado_por':anulado_por
                        }
                        DRS_FILTER.append(new_data)  

        return DRS_FILTER


    def save_excel(self, filename):

        DRS_FILTER = self.get_filtered_data()

        file = openpyxl.load_workbook('./FILTER_OUTPUT/TEMP.xlsx')
        sheet = file.active

        c = 2
        for drs in DRS_FILTER:
            sheet[f'A{c}'] = drs['drs']
            sheet[f'B{c}'] = drs['codigo_modelo']
            sheet[f'C{c}'] = drs['projeto']
            sheet[f'D{c}'] = drs['mercado']
            sheet[f'E{c}'] = drs['cliente']
            sheet[f'F{c}'] = drs['uni_prod']
            sheet[f'G{c}'] = drs['resp_pedido']
            sheet[f'H{c}'] = drs['dep_requerente']
            sheet[f'I{c}'] = drs['data_registo']
            sheet[f'J{c}'] = drs['finalizado_por']
            sheet[f'K{c}'] = drs['finalizado_em']
            sheet[f'L{c}'] = drs['aprovado_por']
            sheet[f'M{c}'] = drs['recusado_por']
            sheet[f'N{c}'] = drs['anulado_por']
            c += 1
        is_folder = os.path.isdir(f'./FILTER_OUTPUT/{filename + '_' + str(date.today())}')
        if is_folder:
            return False
        os.mkdir(f'./FILTER_OUTPUT/{filename + '_' + str(date.today())}')
        file.save(f'./FILTER_OUTPUT/{filename + '_' + str(date.today())}/{filename}.xlsx')
        file.close()
        return 'File saved!'




def main():

    print(r"""
$$$$$$$\  $$$$$$$\   $$$$$$\        $$$$$$$$\ $$$$$$\ $$\    $$$$$$$$\ $$$$$$$$\ $$$$$$$\  
$$  __$$\ $$  __$$\ $$  __$$\       $$  _____|\_$$  _|$$ |   \__$$  __|$$  _____|$$  __$$\ 
$$ |  $$ |$$ |  $$ |$$ /  \__|      $$ |        $$ |  $$ |      $$ |   $$ |      $$ |  $$ |
$$ |  $$ |$$$$$$$  |\$$$$$$\        $$$$$\      $$ |  $$ |      $$ |   $$$$$\    $$$$$$$  |
$$ |  $$ |$$  __$$<  \____$$\       $$  __|     $$ |  $$ |      $$ |   $$  __|   $$  __$$< 
$$ |  $$ |$$ |  $$ |$$\   $$ |      $$ |        $$ |  $$ |      $$ |   $$ |      $$ |  $$ |
$$$$$$$  |$$ |  $$ |\$$$$$$  |      $$ |      $$$$$$\ $$$$$$$$\ $$ |   $$$$$$$$\ $$ |  $$ |
\_______/ \__|  \__| \______/       \__|      \______|\________|\__|   \________|\__|  \__|
"""
    )

    print('\n')

    isGlobal = input('Global search data? (Y/n): ' ).strip().upper()
    if isGlobal == 'Y':
        _global_ = True

        #get years
        years = input('Years to search (separate by comma): ')
        years_list = []
        for y in years.split(','):
            years_list.append(y)


    elif isGlobal == 'N':
        _global_ = False
    else:
        _global_ = True

    open_drs = int(input('Search for closed DRS(1) or open(2) or all (3) (default close): '))
    if open_drs == 2:
        drs_open = True
    elif open_drs == 3:
        drs_open = 'All'
    else:
        drs_open = False

    type_of_filter = int(input('Filter by department(1) or request type(2): ' ))
    
    if type_of_filter == 1:
        dep_filter = input('Department: ').strip().upper()
        request_filter = None

    if type_of_filter == 2:
        request_filter = input('Request type: ').strip().upper()
        dep_filter = None
    
    print('\n')
    print('*'*30)
    print('Search...')
    print('*'*30)
    print('\n')


    x = FILTER_DATA(global_=_global_, request=request_filter, dep=dep_filter, _open=drs_open, filter_year=years_list)
    data = x.get_filtered_data()
    print(f'{len(data)} DRS found!')
    print('\n')
    time.sleep(2)
    
    for j in data:
        print(j)
    
    print('\n')

    if len(data) >= 0:
        print_to_exell = input('Save data in excel? (y/n) :' ).strip().upper()
        if print_to_exell == 'Y':
            while True:
                filename = input('File name: ')
                resp = x.save_excel(filename)
                if not resp:
                    print('Folder already exists')
                else:
                    print('File Saved')
                    break


    



if __name__ == '__main__':
    main()
