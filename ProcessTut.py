# -*- coding: utf-8 -*-
import os
import os.path


def process_01():
    os.chdir('g:\\temp\\lynda')
    for fname in os.listdir():
        mname, exn = os.path.splitext(fname)
        if not os.path.exists(mname):
            os.mkdir(mname)
        os.rename(fname, os.path.join(mname, fname))


def process_02():
    workdir = 'z:\\Incoming\\tt'
    for trtname in os.listdir(workdir):
        orgname = os.path.join(workdir, trtname)
        if os.path.isdir(orgname):
            for subname in os.listdir(orgname):
                if os.path.isdir(os.path.join(orgname, subname)):
                    #newsubname = '[' + trtname + '].' + subname
                    #newsubname = newsubname.replace(' ', '.')
                    #print(os.path.join(orgname, subname), os.path.join(workdir, subname))
                    os.renames(os.path.join(orgname, subname), os.path.join(workdir, subname))
                    break


if __name__ == '__main__':
    process_02()