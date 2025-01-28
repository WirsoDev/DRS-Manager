from app.jiraConn import JiraConn
from app.fileRegistrationDb import Fileregister


#jcon = JiraConn()

data = {
                'drsNumber': '3529',
                'drsVersion': '1',
                'client':'BANAK',
                'tiposPedido': 'IMAGEM PROD, 3D-AMBIENTE, TEAR DOWN',
                'modelName': 'ASTON',
                'modelCode': '1254',
                'tipologia': 'CNT',
                'deadline': 'W20',
            }



if __name__ == "__main__":
    jcon = JiraConn()
    a = jcon.treat_drs(data)
    print(a)
