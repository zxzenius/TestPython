# -*- coding: utf-8 -*-

import os
import os.path
import re


def process(DestFolder):
    for Root, Folders, Files in os.walk(DestFolder, topdown=False):
        for FileName in Files:
            MainName, ExtName = os.path.splitext(FileName)
            if ExtName.lower() in ('.bms', '.bme'):
                NewFileName = FileName
                if FileName[:1] != '#':
                    NewFileName = '#' + FileName
                    #NewFileName= FileName.replace('KURUIZAKE HOMURA NO HANA', '狂イ咲ケ焔ノ華')
                if FileName != NewFileName:
                    FullFileName = os.path.join(Root, FileName)
                    NewFullFileName = os.path.join(Root, NewFileName)
                    os.rename(FullFileName, NewFullFileName)
                    print(FullFileName, NewFileName)


if __name__ == '__main__':
    process('f:\\incoming\\20  Tricoro')



