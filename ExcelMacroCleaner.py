# -*- coding: utf-8 -*-
import win32com.client
import os.path
import os


def Process(FileList, UpVersion=True):
    Total = len(FileList)
    if Total == 0:
        exit();
    Counter = 0
    ProcessedCounter = 0
    print('Processing...')
    xlsApp = win32com.client.Dispatch('Excel.Application')
    xlsApp.Visible = False
    xlsApp.AutomationSecurity = 3  #msoAutomationSecurityForceDisable = 3
    xlsApp.DisplayAlerts = False
    for FileName in FileList:
        Counter += 1
        print('(%d/%d)' % (Counter, Total), FileName, )
        try:
            xlsWorkbook = xlsApp.Workbooks.Open(FileName)
            #xlsWorkbook.ConfilictResolution = 2 #xlLocalSessionChanges = 2
            xlsNames = xlsWorkbook.Names
            for n in xlsNames:
                if (not n.Visible) and n.MacroType == 2:
                    n.Delete()
            if UpVersion:
                NewFileName = NewName(FileName)
                #xlsFormat = win32com.client.constants.xlOpenXMLWorkbook  #xlOpenXMLWorkbook  = 51
                xlsWorkbook.SaveAs(NewFileName, 51)
                xlsWorkbook.Close()
                os.remove(FileName)
            else:
                NewFileName = FileName
                xlsWorkbook.Save()
                xlsWorkbook.Close()
            ProcessedCounter += 1
        except:
            print('Failed', end='\n')
        finally:
            print('->', NewFileName, end='\n')
    print('%d files have been processed.' % ProcessedCounter)


def NewName(FullFileName):
    ExtStr = '.xlsx'
    Path, FileName = os.path.split(FullFileName)
    MainName, ExtName = os.path.splitext(FullFileName)
    NewFileName = os.path.join(Path, ''.join((MainName, ExtStr)))
    Suffix = 1
    while os.path.exists(NewFileName):
        NewFileName = os.path.join(Path, ''.join((MainName, '.', str(Suffix), ExtStr)))
        Suffix += 1
    return NewFileName


def Start(TargetPath):
    TargetExtStr = '.xls'
    xlsFileList = []
    print('Searching...')
    for Root, Dirs, Files in os.walk(TargetPath):
        for File in Files:
            FileName, ExtName = os.path.splitext(File)
            if ExtName == TargetExtStr.lower():
                FullFileName = os.path.join(Root, File)
                xlsFileList.append(FullFileName)
    print(str(len(xlsFileList)), 'files have been found.')
    Process(xlsFileList)


def KillK4():
    TargetPath = os.path.expandvars('$appdata\\Microsoft\\Excel\\XLSTART')  #For K4.xls
    for Item in os.listdir(TargetPath):
        File = os.path.join(TargetPath, Item)
        if os.path.isfile(File):
            try:
                os.remove(File)
            except:
                print("Found K4.xls, but i can't kill it.")
            else:
                print('K4.xls has been killed')


if __name__ == '__main__':
    #KillK4()
    #Start(os.getcwd())
    Start('D:\\Work\\project\\YunNan.QuJing.COG.LNG.200K\\Equip\\MR.MakeUp')
    #Start('d:\\work\\dev\macrovirus')
    input('Press Enter to quit')

