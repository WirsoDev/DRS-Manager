import shutil
import os
from dotenv import load_dotenv
from datetime import date

load_dotenv()


class BACK_UP:
    def __init__(self):
        self.dbfile = os.environ.get('DIRDATABASE')
        self.backup_dir = os.environ.get('BACKUPS')
    

    def mk_backup(self):
        
        _date = date.today()
        filename = f'backup_DRSDB_{_date}.xlsx'
        filename_complete = f'{self.backup_dir}/{filename}'

        #check if file alreay saved
        allfiles = os.listdir(self.backup_dir)
        if filename in allfiles:
            return
       
        try:
            shutil.copy(self.dbfile, filename_complete)
            print('** Make DB backup **')
            return True
        except:
            print('** Some problem making DB backup **')
            return False
        

        

