from app.fileRegistrationDb import Fileregister
from app.margePdfFiles import MargePdfFiles
import time
from datetime import date

# main program 
def main():
    print(f'Program running: {date.today()}!')
    while True:
        # db manager
        try:
            manageDb = Fileregister()
            manageDb.lastDrsCreated()
            manageDb.registFiles()
            managePdf = MargePdfFiles()
            managePdf.MargeDuplicatedFiles()
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



