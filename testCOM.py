# -*- coding: utf-8 -*-
import win32com.client
import os.path

#For testing, reading and writing ExcelFile
def LinkExcel():
    #Set targetFile
    xlsFile = 'E:\\My.Dev\\TestRes\\test1.xlsx'
    if not os.path.exists(xlsFile):
        exit()
    #Create Link
    xlsApp = win32com.client.Dispatch('Excel.Application')
    xlsWorkbook = xlsApp.Workbooks.Open(xlsFile)
    xlsSheets = xlsWorkbook.Sheets
    xlsSheet = xlsSheets[0]
    #Used Zone
    SourceZone = xlsSheet.UsedRange
    #Extract Data of target zone
    for Cell in SourceZone:
        print(Cell.Address, Cell.Value, sep=': ')

    #Close File & Process
    xlsWorkbook.Close()
    xlsApp.Quit()
    xlsApp = None


if __name__ == '__main__':
    LinkExcel()
