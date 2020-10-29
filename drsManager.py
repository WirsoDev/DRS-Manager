from app.fileRegistrationDb import Fileregister
from app.margePdfFiles import MargePdfFiles
import time
from datetime import date

# main program 
def main():
    print(f'Program running at: {date.today()}!!')
    while True:
        try:
            manageDb = Fileregister()
            manageDb.registFiles()
            time.sleep(10)
        except PermissionError:
            print('File open!')
            time.sleep(10)
            pass
        except:
            print(f'ERROR! at: {date.today()}')


