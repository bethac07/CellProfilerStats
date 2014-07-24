import os

count=0
musicdir=r'C:\Users\Beth Cimini\Music'
for i in os.listdir(musicdir):
    newdir=os.path.join(musicdir,i)
    if os.path.isdir(newdir):
        dirlist1=os.listdir(newdir)
        #print dirlist1
        if len(dirlist1)>0:
            for j in dirlist1:
                #print j
                #print os.path.join(newdir,j)
                subdir=os.path.join(newdir,j)
                if os.path.isdir(subdir):
                    songs=os.listdir(subdir)
                    for k in songs:
                        if '1.m4p' in k:
                            count+=1
                            os.remove(os.path.join(subdir,k))
print count
