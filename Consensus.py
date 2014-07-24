# -*- coding: utf-8 -*-
"""
Created on Wed Apr 13 14:20:07 2011
Designed to take the kinds of sheets made by OutputSigsXL and scan them for things that are consensus-
still in progress
@author: Beth Cimini
"""

import xlrd
import easygui as eg

#pull headings
whichfilein=eg.fileopenbox()
print whichfilein
filein=xlrd.open_workbook(whichfilein)
a=[]
for sheet in filein.sheets():
    a.append(sheet.row(0))
#print a[0][1].value[8:-4]

#create map of where identical headings are
b={}
for i in range(len(a)):
    for j in range(1,len(a[i])):
        if a[i][j].value[8:-4] not in b:
            b[a[i][j].value[8:-4]]=[(i,j)]
        else:
            b[a[i][j].value[8:-4]].append((i,j))
#print b

#test if values are the same
for i in b.keys():
      if len(b[i])>1:
          print ''
          print i, 'n=',len(b[i])
          for j in range(1,filein.sheet_by_index(b[i][0][0]).nrows):
              z=[]
              for y in b[i]:
                  if len(filein.sheet_by_index(y[0]).cell(j,y[1]).value)>1:
                      z.append(filein.sheet_by_index(y[0]).cell(j,y[1]).value[-1])
                  else:
                      z.append('')
              m=z[0]
              if m=='':
                  continue
              identical=True
              for n in z:
                  if n!=m:
                      identical=False
              if identical==True:
                  print filein.sheet_by_index(y[0]).cell(j,0).value, 'significantly goes',m
                  