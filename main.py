import PySimpleGUI as sg
import os
import xlsxwriter

def DirCheck(folder):
    return True if folder and os.path.isdir(folder) else False

def CreateFolder(folder="", foldercount=1, foldername='Report'):
    reportfolder=os.path.join(folder,foldername+'_'+str(foldercount))
    try:
        if not DirCheck(reportfolder):
            os.makedirs(reportfolder)
            print (reportfolder)
            return reportfolder
        else:
            '''return reportfolder if [f for f in os.listdir(reportfolder) if not f.startswith('.')] == [] else '''
            CreateFolder(folder ,foldercount=int(foldercount)+1)
    except OSError:
        print ('Error: Creating directory. ' +  foldercount)


folder=sg.PopupGetFolder('Find folder','Browse')
print(folder)
files=[f for f in os.listdir(folder) if f.endswith('.mp4') or f.endswith('.MP4') ]
print (files)
for i in files:
    print (os.path.abspath(os.path.join(folder,i)))

reportfolder = CreateFolder(folder)
print (reportfolder)
name='/Report.xlsx'
path = os.path.join(reportfolder,name)
print (path)
wb=xlsxwriter.Workbook(path)
