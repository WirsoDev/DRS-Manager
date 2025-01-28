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
        self.reports = REPORTS(global_=self.isGobal, filter_year=self.filter_year)
        self.data = self.reports.getAllDrsData()
        self.dep = dep
        self.request = request
        self.isopen = _open


    def requests_filter(self):
        requests = {
            'MRK': [
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
            'CUSTEIO': [
                'CUSTEIO',
                'REVISÃO DE PREÇO'
            ],
            'DID': [
                'CERTIFICAÇÃO',
                'PRODUÇÃO DE AMOSTRA',
                'PROTOTIPO',
                'PROTOTIPO CUSTEIO',
                'PROTOTIPO + CERTIFICAÇÃO',
                'CERTIFICAÇAO',
                'REENGENHARIA',
                'CUSTEIO'
            ]
        }

        if self.request:
            return [self.request]

        if self.dep:
            return requests.get(self.dep, [])  # Evita erro caso a chave não exista

        # Combina todas as categorias em uma única lista
        return [value for sublist in requests.values() for value in sublist]




    def get_filtered_data(self):
        data = self.data
        accepted_requests = self.requests_filter()
        DRS_FILTER = []

        # Normalize os accepted_requests para comparação
        accepted_requests = [req.strip().upper() for req in accepted_requests]

        for drs_id, drs_data in data.items():
            # Extrair informações básicas
            drs_n = drs_data.get('drs_n')
            cliente = drs_data.get('cliente')
            codigo_modelo = drs_data.get('codigo_modelo')
            mercado = drs_data.get('mercado')
            Uni_prod = drs_data.get('Uni_prod')
            Resp_pedido = drs_data.get('Resp_pedido')
            dep_requerente = drs_data.get('dep_requerente')
            finalizado_por = drs_data.get('finalizado_por')
            recusado_por = drs_data.get('recusado_por')
            anulado_por = drs_data.get('anulado_por')
            aprovado_por = drs_data.get('aprovado_por')

            # Tratamento de datas
            if self.isopen is True:
                data_registo = drs_data.get('data_registo')
                new_date = parse_date(data_registo)
                new_date_finalizado = None
            elif self.isopen is False:
                data_registo = drs_data.get('finalizado_em')
                new_date = parse_date(data_registo)
                new_date_finalizado = new_date
            else:  # Caso seja "All"
                data_registo = drs_data.get('data_registo')
                data_finalizado = drs_data.get('finalizado_em')
                new_date = parse_date(data_registo)
                new_date_finalizado = parse_date(data_finalizado)

            # Tipos de pedidos com seus índices
            tipos_pedidos = [
                (str(drs_data.get('tipo_pedido_1', '')).strip().upper(), 1),
                (str(drs_data.get('tipo_pedido_2', '')).strip().upper(), 2),
                (str(drs_data.get('tipo_pedido_3', '')).strip().upper(), 3),
            ]

            # Para cada tipo de pedido válido, cria uma entrada separada
            for tipo_pedido, pedido_num in tipos_pedidos:
                if tipo_pedido and tipo_pedido in accepted_requests:
                    print(f"DRS {drs_n} - Incluindo tipo de pedido {pedido_num}: {tipo_pedido}")
                    new_data = {
                        'drs': drs_n,
                        'codigo_modelo': codigo_modelo,
                        'projeto': tipo_pedido,  # Usa o tipo de pedido específico
                        'mercado': mercado,
                        'cliente': cliente,
                        'uni_prod': Uni_prod,
                        'resp_pedido': Resp_pedido,
                        'dep_requerente': dep_requerente,
                        'data_registo': new_date,
                        'finalizado_por': finalizado_por,
                        'finalizado_em': new_date_finalizado,
                        'aprovado_por': aprovado_por,
                        'recusado_por': recusado_por,
                        'anulado_por': anulado_por,
                        'pedido_num': pedido_num  # Opcional: adiciona o número do tipo de pedido
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


def parse_date(date_string):
    if date_string is None or date_string == "None":  # Handle both None and string 'None'
        return None
    if isinstance(date_string, str):
        for fmt in ('%d/%m/%Y', '%Y-%m-%d %H:%M:%S'):
            try:
                return datetime.strptime(date_string, fmt)
            except ValueError:
                continue
        raise ValueError(f"Date format not recognized: {date_string}")
    if isinstance(date_string, datetime):  # Already a datetime object
        return date_string
    raise ValueError(f"Unsupported date type: {type(date_string)}")



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
        years = input('Year to search: ')


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


    x = FILTER_DATA(global_=_global_, request=request_filter, dep=dep_filter, _open=drs_open, filter_year=years)
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
