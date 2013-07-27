# -*- coding: utf-8 -*-

import os
import os.path
import re

__author__ = 'zxz'

#"E:\Soft\mediatool\mkvtoolnix\mkvmerge.exe" -o "F:\\Backup\\Drive.G\\爱情公寓.Ipart.S02\\爱情公寓.Ipartment.S02E01.HDTV.UNCUT.MiniSD-TLF (1).mka"  "--language" "1:chi" "--track-name" "1:普通话" "--default-track" "1:yes" "--forced-track" "1:no" "-a" "1" "-D" "-S" "-T" "--no-global-tags" "--no-chapters" "(" "F:\\Backup\\Drive.G\\爱情公寓.Ipart.S02\\爱情公寓.Ipartment.S02E01.HDTV.UNCUT.MiniSD-TLF.mkv" ")" "--track-order" "0:1"

def GetCMD(SourceFile, DestFile):
    mkvmergePATH='"E:\\Soft\\mediatool\\mkvtoolnix\\mkvmerge.exe"'
    return(mkvmergePATH+' -o "{1}"  "--language" "1:chi" "--track-name" "1:普通话" "--default-track" "1:yes" ' \
                        '"--forced-track" "1:no" "-a" "1" "-D" "-S" "-T" "--no-global-tags" "--no-chapters" "(" "{0}"' \
                        ' ")" "--track-order" "0:1"'.format(SourceFile, DestFile))

def MKVAudioSplte(SourcePath, DestPath, CMDFile):
    os.chdir(SourcePath)
    CmdFile=open(CMDFile, mode='w')
    for MKVFile in os.listdir():
        if os.path.isfile(MKVFile):
            FileName=os.path.basename(MKVFile)
            MainName, ExtName=os.path.splitext(FileName)
            if ExtName != '.mkv':
                continue
            DestFile=DestPath + '爱情公寓.Ipartment' + re.search('\.S\w+\.\w+',MainName).group() + '.mka'
            CmdFile.write(GetCMD(os.path.join(SourcePath, MKVFile), DestFile)+'\n')
            #print(DestFile)




if __name__ == '__main__':
    TargetPath='F:\\Backup\\Drive.G\\爱情公寓.Ipart.S03\\'
    MKVAudioSplte(TargetPath, TargetPath, TargetPath + 'MKVSplit.bat')

