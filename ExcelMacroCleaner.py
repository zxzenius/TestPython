# -*- coding: utf-8 -*-
import win32com.client
import os.path
import os


def Process(FileName):
    print(FileName, )
    xlsApp = win32com.client.Dispatch('Excel.Application')
    xlsApp.AutomationSecurity = 3  #msoAutomationSecurityForceDisable = 3
    xlsApp.DisplayAlerts = False
    xlsWorkbook = xlsApp.Workbooks.Open(FileName)
    #xlsWorkbook.ConfilictResolution = 2 #xlLocalSessionChanges = 2
    xlsNames = xlsWorkbook.Names
    for n in xlsNames:
        if n.Visible == False:
            n.Delete()
    NewFileName = NewName(FileName)
    #xlsFormat = win32com.client.constants.xlOpenXMLWorkbook  #xlOpenXMLWorkbook  = 51
    xlsWorkbook.SaveAs(NewFileName, 51)
    xlsWorkbook.Close()
    print('->', NewFileName, end='\n')
    return (True)


def NewName(FullFileName):
    ExtStr = '.xlsx'
    Path, FileName = os.path.split(FullFileName)
    MainName, ExtName = os.path.splitext(FullFileName)
    NewFileName = os.path.join(Path, MainName + ExtStr)
    Suffix = 1
    while os.path.exists(NewFileName):
        NewFileName = os.path.join(Path, MainName + '.' + str(Suffix) + ExtStr)
        Suffix = Suffix + 1
    return (NewFileName)


def Start(TargetPath):
    TargetExtStr = '.xls'
    xlsFileList = []
    counter = 0
    print('Searching...')
    for Root, Dirs, Files in os.walk(TargetPath):
        for File in Files:
            FileName, ExtName = os.path.splitext(File)
            if ExtName == TargetExtStr.lower():
                #counter = counter + 1
                FullFileName = os.path.join(Root, File)
                xlsFileList.append(FullFileName)
                # Process(FullFileName)
                #print(FileFullName, end='\n')
    print(str(len(xlsFileList)), 'files have been found.')
    print('Processing...')
    for xlsFile in xlsFileList:
        if Process(xlsFile):
            counter = counter + 1
    print(str(counter), 'files have been processed.')


if __name__ == '__main__':
    Start(os.getcwd())
    #Start('d:\\work\\dev\macrovirus')
    input('Press Enter to quit')

