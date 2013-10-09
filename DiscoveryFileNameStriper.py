import os
import os.path

TargetFolder = 'g:\\discovery'

os.chdir(TargetFolder)

for filename in os.listdir():
    if os.path.isdir(filename):
        continue
    newname = filename.replace('Discovery：', '')
    newname = newname.replace('国家地理频道.', '')
    lpos = newname.find('(')
    #if lpos > -1:
    #newname = newname[:lpos] + newname[newname.find(')')+1:]
    #print(newname)
    os.rename(filename, newname)
    

        