import os
import os.path
import re
import FileArranger

import sys


def WavName(cuefile):
    for line in open(cuefile, 'rb'):
        #print(line)
        match = re.search(b'(?<=FILE) +"(.+\.ape)"', line)
        if match:
            #print(cuefile, match.group(1).decode())
            try:
                return(match.group(1).decode())
            except UnicodeDecodeError:
                return(match.group(1).decode('gbk'))

def iscue(filename):
    return(os.path.splitext(filename)[1].lower()=='.cue')
        
def process(workpath):   
    FileArranger.process(workpath)
    counter = 0
    for path in os.listdir():
        #print('path:',path)
        try:
            os.chdir(path)
            for file in os.listdir():
                #print(file)
                mname, extname = os.path.splitext(file)
                if extname.lower() == '.cue':
                    apename = WavName(file)
                    apefile = mname + '.ape'
                    if (not os.path.exists(apename)) and (os.path.exists(apefile)):
                        os.rename(apefile,  apename)
                        counter += 1                        
                        break
            os.chdir('..')   
        except:
            print('err: ',path, sys.exc_info()[1])
            continue
    print('%s processed'%counter)       
            
def patharrange(workpath):
    os.chdir(workpath)
    for path in os.listdir():
        match = re.search('(?<=volume).*(\d+)\).+CD(\d+)', path, flags=re.IGNORECASE)
        #print(re.split('\w+', path))
        if match:
            #print(workpath,match.groups())
            mpath, spath = ('%02d'%int(d) for d in match.groups())
            mpath = 'Volume' + mpath
            spath = 'CD' + spath
            #print(path, os.path.join(mpath, spath))
            os.renames(path, os.path.join(mpath, spath)) 
                 
 

             

if __name__ == '__main__':
    workpath = ('f:\\mozart', 'f:\\bach')
    #process('f:\\mozart')
    for path in workpath:
        process(path)
        patharrange(path)
            

        
