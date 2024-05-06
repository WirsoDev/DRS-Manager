from app.fileRegistrationDb import Fileregister
from app.margePdfFiles import MargePdfFiles
from app.backups import BACK_UP
import time
from datetime import date
from registStatus import registStatus

# main program 
def main():
    cicle = 1
    print(f'++++++++++++++++++++++++++++++++')
    print(f'Program running: {date.today()}!')
    print(f'++++++++++++++++++++++++++++++++')
    while True:
        st = time.time()
        print(f' -> RUNING CICLE: {cicle}')
        # db manager
        try:
            manageDb = Fileregister()
            manageDb.lastDrsCreated()
            manageDb.registFiles()
            registStatus()

            #manage backups
            bk_manager = BACK_UP()
            bk_manager.mk_backup()

            time.sleep(10)
        except PermissionError:
            print('ERROR: File open!')
            time.sleep(10)
            pass
        except KeyboardInterrupt:
            print('By :)\nExiting...')
            time.sleep(2)
            break
        except:
            print(f'ERROR! at: {date.today()}')
        ft = time.time()
        print(f' -> CLOSE CICLE: {cicle} : Running time: {ft - st} s')
        cicle += 1


