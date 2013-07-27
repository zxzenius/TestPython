import os
import os.path

os.chdir('g:\\temp\\lynda')
for fname in os.listdir():
    mname, exn = os.path.splitext(fname)
    if not os.path.exists(mname):
        os.mkdir(mname)
    os.rename(fname, os.path.join(mname, fname))
    
    