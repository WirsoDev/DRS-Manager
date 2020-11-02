import PyPDF2
import os
from dotenv import load_dotenv
import PyPDF2 

load_dotenv()


class MargePdfFiles:

    def __init__(self):
        self.path = os.environ.get('DIRPDFFILES')
        self.files = os.listdir(self.path) # array of existing files


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
            return 'Folder is empty!'

    def MargeDuplicatedFiles(self):
        '''Marge the files in versions -> PDF FILE'''
        print(self.Duplicatefilesfinder())
        if self.Duplicatefilesfinder():
            marquer = self.Duplicatefilesfinder()[0].split('_')[1].strip()
            bucketFiles = []
            for file in self.Duplicatefilesfinder():
                fileNumber = ''.join(file.split('_')[1]).strip()
                if fileNumber == marquer or len(bucketFiles) == 0:
                    bucketFiles.append(file)
            
            print(bucketFiles)

                

