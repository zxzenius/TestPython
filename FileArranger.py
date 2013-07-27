import os
import os.path

def arrange(fname, pathname = '', newfname = ''):
    if not pathname:
        pathname, exn = os.path.splitext(fname)
    if not os.path.exists(pathname):
        os.mkdir(pathname)
    if not newfname:
        newfname = fname
    os.rename(fname, os.path.join(pathname, newfname))
    
def process(workpath):
    os.chdir(workpath)
    for fname in os.listdir():
        if os.path.isfile(fname):
            arrange(fname)