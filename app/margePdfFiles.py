import PyPDF2
import os
from dotenv import load_dotenv
import PyPDF2
from shutil import copyfile
from datetime import datetime

load_dotenv()


class MargePdfFiles:

    def __init__(self):
        self.path = os.environ.get('DIRPDFFILES')
        self.backup = os.environ.get('BACKPDF')
        self.files = []
        for pdf in os.listdir(self.path): 
            if '.pdf' in pdf:
                self.files.append(pdf)


    def Duplicatefilesfinder(self):
        '''Return a list of duplicate files in folder -> ARRAY'''
        if self.files:
            listoffiles = []
            for files in self.files:
                breakfile = files.split('_')
                fileNumber = ''.join(breakfile[1].split())
                listoffiles.append(fileNumber)

            duplicatedFile = [x for n, x in enumerate(listoffiles) if x in listoffiles[:n]]
            duplicatedDRS = []

            for file in self. files:
                edfile = file.split('_')
                edfile1 = ''.join(edfile[1].split())
                if edfile1 in duplicatedFile:
                    duplicatedDRS.append(file) 
            duplicatedDRS.sort()
            return duplicatedDRS
        else:
            return False

    def MargeDuplicatedFiles(self):
        '''Marge the files in versions -> PDF FILE'''
        if self.Duplicatefilesfinder():
            # bucket all the same DRS files
            # save duplicates to backup folder
            marquer = self.Duplicatefilesfinder()[0].split('_')[1].strip()
            bucketFiles = []
            for file in self.Duplicatefilesfinder():
                fileNumber = ''.join(file.split('_')[1]).strip()
                if fileNumber == marquer or len(bucketFiles) == 0:
                    bucketFiles.insert(0,file)
                    copyfile(f'{self.path}/{file}', f'{self.backup}/{file}')

            # create a new pdf file
            # marge all pdf to the new file
            # Remove duplicates from folder
            margedFile = PyPDF2.PdfFileMerger()
            for x in range(len(bucketFiles)):
                margedFile.append(PyPDF2.PdfFileReader(f'{self.path}/{bucketFiles[x]}', 'rb'))
                os.remove(f'{self.path}/{bucketFiles[x]}')

            # save new file -> name : DRS_NUMBER_NUMBER OF VERSION_.pdf
            numberOfEd = f'{len(bucketFiles)}'
            MargedDrsName = f'DRS_{marquer}_ED_{numberOfEd}'
            margedFile.write(f'{self.path}/{MargedDrsName}.pdf')

            now = datetime.now()
            print(f'New files marged! At: {now}')
        else:
            return False
                

