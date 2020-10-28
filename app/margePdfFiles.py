'''
Check for files in folder
if finds DRS with the same number, marge them in a single file

Test sort in duplicates
Try marge by DRS Number
'''

import PyPDF2
import os
import dotenv

class MargePdfFiles:

    def __init__(self):
        self.filepath = os.environ.get('DIRPDFFILES')
        self.files = os.listdir(self.filepath) # array of existing files
        print('all files in folder:', self.files, '\n')

    
    def Duplicatefilesfinder(self):
        '''Return a list of duplicate files in folder -> ARRAY'''
        listoffiles = []
        for files in self.files:
            breakfile = files.split('_')
            fileNumber = ''.join(breakfile[1].split())
            listoffiles.append(fileNumber)

        duplicatedFile = [x for n, x in enumerate(listoffiles) if x in listoffiles[:n]]
        duplicatedDRS = []

        for file in self.files:
            edfile = file.split('_')
            edfile1 = ''.join(edfile[1].split())
            if edfile1 in duplicatedFile:
                duplicatedDRS.append(file) 
        duplicatedDRS.sort()
        return duplicatedDRS

    def MargeDuplicatedFiles(self):
        '''Marge the files in versions -> PDF FILE'''
        pass
