import os
import os.path

os.chdir("f:\\temp\\ForFILEPROCESS")

for pathstr in os.listdir(os.getcwd()):
    if os.path.isdir(pathstr):
        continue
    mainname, extname = os.path.splitext(pathstr)
    mainname = mainname.replace('.文字版', '')
    if not os.path.isdir(mainname):
        os.mkdir(mainname)
    os.rename(pathstr, '\\'.join((mainname, pathstr.replace('.文字版', ''))))


    

    
