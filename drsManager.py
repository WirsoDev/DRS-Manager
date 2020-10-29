# main file

from app.fileRegistrationDb import Fileregister
from app.margePdfFiles import MargePdfFiles

import time


# main program 
def main():
    print('Program are running!!')
    while True:
        try:
            pdfmanager = MargePdfFiles().Duplicatefilesfinder()
            time.sleep(10)
        except:
            print('Something go wrong!')
            break


