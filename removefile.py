import os
import sys
from os import path

def removefile(keyfilelist, originpath):
    for originpath,dirs,files in os.walk(originpath):
        for filename in files:
            if keyfile in filename:
                print("remove file: ", filename)
                os.remove(originpath+'/'+filename)


def findfile(keyword,currentPath):
    filelist=[]
    keyfilelist=[]
    for currentPath,dirs,files in os.walk(currentPath):
        for name in files:                
            fitfile=filelist.append(os.path.join(currentPath, name))   
    for i in filelist:            
        if os.path.isfile(i):
            if keyword in os.path.split(i)[1]:
		keyfilewithdot = os.path.basename(os.path.split(i)[1])
		keyfile = keyfilewithdot.split('.')[0]
		keyfilelist.append(keyfile)
    return keyfilelist


if __name__ == '__main__':
    abspath     = sys.path[0]
    currentpath = sys.argv[2]
    originpath  = sys.argv[3]
    buildnumber = sys.argv[4]
    print("abspath: ", abspath)
    print("currentpath :", currentpath)
    print("originpath :", originpath)
    print("buildnumber :", buildnumber)
    keyfilelist = findfile(buildnumber, abspath+"/"+currentpath)
    print("keyfile :", keyfilelist)
    for keyfile in keyfilelist:
        removefile(keyfile, originpath)
