# -*- coding: utf-8 -*-
"""
Created on Thu Apr 26 16:14:48 2012

@author: Beth Cimini
"""

def readinput(m):
    n=[]
    first=None
    comma=-1
    for i in range(len(m)):
        if m[i]==',':
            try:
                n.append(int(m[i-1]))
            except:
                pass
            comma=i
        if m[i]=='-':
            if first==None:
                first=i
                start=int(m[comma+1:first])    
            else:
                second=i
                stop=int(m[first+1:second])
                interval=int(m[second+1:])
    for i in range(start,stop,interval):
        n.append(i)
    startframe=(start-1)/2
    inter=interval-2
    n.sort()
    r=str(n)[1:-1]
    for i in range(len(r)-1,0,-1):
        if r[i]==' ':
            r=r[:i]+r[i+1:]        
    return r,startframe,inter
    
print readinput('1,3-76-6')