from app.fileRegistrationDb import Fileregister
from app.margePdfFiles import MargePdfFiles
import time
from datetime import date
from registStatus import registStatus

# main program 
def main():
    st = time.time()
    cicle = 1
    print(f'++++++++++++++++++++++++++++++++')
    print(f'Program running: {date.today()}!')
    print(f'++++++++++++++++++++++++++++++++')
    while True:
        print(f'     - RUNING CICLE: {cicle}')
        # db manager
        try:
            manageDb = Fileregister()
            manageDb.lastDrsCreated()
            manageDb.registFiles()
            registStatus()
            time.sleep(10)
        except PermissionError:
            print('File open!')
            time.sleep(10)
            pass
        except KeyboardInterrupt:
            print('By :)\nExiting...')
            time.sleep(2)
            break
        except:
            print(f'ERROR! at: {date.today()}')
        ft = time.time()
        print(f'     - CLOSE CICLE: {cicle} : Running time: {ft - st}')
        cicle += 1


