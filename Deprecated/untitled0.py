# -*- coding: utf-8 -*-
"""
Created on Tue Mar 01 16:41:44 2011

@author: Beth Cimini
"""
#from PIL import Image
import xlrd
import xlwt
from xlutils.copy import copy
import os
import easygui as eg
import csv
import HandyXLModules as HXM

class spots(list):
    def __init__(self,book,sheet,minlength=0):
        book= xlrd.open_workbook(book)
        sh = book.sheet_by_name(sheet)
        headings=sh.row_values(0)
        self.append(headings)
        for i in range(len(headings)):
            if 'Label' in headings[i]:
                labelcolumn=i
        alllabels=sh.col_values(labelcolumn)
        maxlabel=int(max(alllabels[1:]))
        for j in range(1,(maxlabel+1)):
            if alllabels.count(j)>=minlength:
                a=[]
                for k in range(len(alllabels)):
                    if alllabels[k]==j:
                        a+=[sh.row_values(k)]
                self.append(a)

    def xyandt(self):
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        a=[]
        for i in self[1:]:
            b=[i[0][0]]
            for j in i:
                b.append((j[x],j[y]))
            a.append(b)
        return a
        
    def xyandtindiv(self,indiv):
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        a=[self[indiv][0][0]]
        for j in self[indiv]:
            a.append((j[x],j[y]))
        return a
        
    

def combinecsv(csvin): #Combines all .csv files into a single .xls file
    w=xlwt.Workbook() #Create new excel file
    folder,filelist=os.path.split(csvin)
    c=len(filelist) #see how long the filename is- for removing the .csv extension below
    while c>=35: #Sheet names can only be 31 characters long- prompt user if shorter sheet name required
        filelist=(eg.enterbox('Filename '+filelist+' is too long- enter a shorter one')+'.csv')
        c=len(filelist)
    b=w.add_sheet(filelist[:-4]) #name each sheet according to the name of it's .csv file
    a=csv.reader(open(csvin)) #Read the file
    t=[]
    for rows in a:
        t.append(rows) #Transfer CSV's into python lists for use below
    for i in range(len(t)): #for all rows of information
        for j in range(len(t[0])): #for each piece of data in the row
            try:
                z=float (t[i][j]) #If the values are numbers, format them as such
                t[i][j]=z
            except:
                pass
            b.write(i,j,t[i][j]) #write each piece of data to it's own sheet
    k=os.path.join(folder,filelist[:-3]+'xls') #Save the excel file
    w.save(k)
    return k,filelist[:-4]


def runCSVIO():    
    a=eg.fileopenbox(msg='Choose input CSV file',default='*.csv') #select a file, open it
    if os.path.isfile(a[:-3]+'xls'):
        filename=a[:-3]+'xls'
        sheetname=os.path.split(a)[1][:-4]
    else:
        filename,sheetname=combinecsv(a)
    bookname=a[:-3]+'xls'
    reuse=eg.indexbox('Where to save the output?',choices=['Add to a new sheet in the input file',
                                                           'Create a new file', 
                                                           'Add to a new sheet in another file',])
    if reuse==0: #If the user chooses to add a new sheet to the existing file
        book=xlrd.open_workbook(bookname)
        writebook=copy(book) #copy to xlutils to make wrtable
        newsheetname=eg.enterbox('What do you want to call the new sheet?')
        writesheet=writebook.add_sheet(newsheetname) #add the new sheet
    if reuse==1: #If the user wants to make a new file
        w=eg.filesavebox(msg='What do you want to name the file?',filetypes=["*.xls"])+'.xls'
        writebook=xlwt.Workbook() #make a new file
        newsheetname=eg.enterbox('What do you want to call the new sheet?')
        writesheet=writebook.add_sheet(newsheetname) #give it a sheet
    if reuse==2: #If the user wants a new sheet in an old file
        o=eg.fileopenbox(msg='Choose Excel file to write to',default='*.xls')
        otherbook=xlrd.open_workbook(o) #Open it
        writebook=copy(otherbook) #make it writable
        newsheetname=eg.enterbox('What do you want to call the new sheet?')
        writesheet=writebook.add_sheet(newsheetname) #add the new sheet

    gate=eg.boolbox(msg='Do you want to discard tracks shorter than a given length?', choices=('Yes','No'))
    if gate:
        hasint=False
        setgate=eg.enterbox(msg='What is the minimum track length to consider?')
        while hasint==False:
            try:
                setgate=int(setgate)
                hasint=True
            except:
                setgate=eg.enterbox(msg='What is the minimum track length to consider?')
    else:
        setgate=0

    writesheet.write(0,0,'Spot Number')
    writesheet.write(0,1,'Frame Number')
    writesheet.write(0,2, 'X coordinate')
    writesheet.write(0,3, 'Y coordinate')
        
    e=spots(filename,sheetname,setgate)
    f=e.xyandt()
    writesheetrows=2
    for i in range(1,len(f)):
        writesheet.write(writesheetrows,0,i)
        HXM.writesheet(writesheet,f[i][1:],writesheetrows,2)
        HXM.writecol(writesheet,range(len(f[i])-1),1,writesheetrows)
        writesheetrows+=len(f[i])
        
    if reuse==0: #if we're saving the input file, save it with the input name
        writebook.save(bookname)
    elif reuse==1: #if we're saving a new file, save it with the name the user input
        writebook.save(w)
    elif reuse==2: #if we're saving an old file, save it with it's original name
        writebook.save(o)

#if __name__=='__main__':
 #   runCSVIO()

from PIL import Image, ImageDraw, ImageFont


#font=ImageFont.truetype('C:\Windows\Fonts\BKANT.TTF',10)

a=r'D:\ImageTesterOut\DefaultOUT_TRF1Telomeres.xls'
b=r'DefaultOUT_TRF1Telomeres'

e=spots(a,b,50)
print len(e)

colorlist=['black','blue','crimson','darkorange','darkorchid','darkcyan','burlywood','deeppink','lawngreen','indigo','seashell','orangered','olive','powderblue','yellow']

f=e.xyandt()
"""linedict={}
for i in range(len(f)):
    for j in range(1,len(f[i])-1):
        for z in range(int(j+f[i][0]),len(f[i])-1+int(f[i][0])):
            if z in linedict.keys():
                #linedict[z].append(((f[i][j],f[i][j+1]),colorlist[i%16]))
                linedict[z].append((f[i][j],colorlist[i%16]))
            else:
                #linedict[z]=[((f[i][j],f[i][j+1]),colorlist[i%16])]
                linedict[z]=[(f[i][j],colorlist[i%16])]"""
                
"""for j in range(1,len(i)-1):
        draw.line((i[j],i[j+1]),fill='red')
    if i[0]==1:
        for j in range(1,len(i)-1):
            draw.line((i[j],i[j+1]),fill='blue')
    draw.text((i[1][0]-5,i[1][1]-5),str(f.index(i)),fill='white',font=font)"""
#print linedict
"""for i in range(max(linedict.keys())):
    c=r'D:\ImageTesterOut\47A_03_R3D_D3D_PRJ-1_'+str(i)+'_TRF1.bmp'
    dd=Image.open(c)
    d=dd.convert('RGB')
    #import ImageDraw,ImageFont
    draw=ImageDraw.Draw(d)
    if i in linedict.keys():
        for j in linedict[i]:
            #print j
            makecircle=[(j[0][0]+3,j[0][1]+0),(j[0][0]+3,j[0][1]+1),(j[0][0]+2,j[0][1]+2),(j[0][0]+1,j[0][1]+3),(j[0][0]+0,j[0][1]+3),(j[0][0]-1,j[0][1]+3),(j[0][0]-2,j[0][1]+2),(j[0][0]-3,j[0][1]+1),(j[0][0]-3,j[0][1]+0),(j[0][0]-3,j[0][1]-1),(j[0][0]-2,j[0][1]-2),(j[0][0]-1,j[0][1]-3),(j[0][0]-0,j[0][1]-3),(j[0][0]+1,j[0][1]-3),(j[0][0]+2,j[0][1]-2),(j[0][0]+3,j[0][1]-1)]
            draw.polygon(makecircle,outline=j[1])
            #draw.line(j[0],fill=j[1],width=5)
    del draw
    d.save(r'D:\ImageTesterOut\47A_03_R3D_D3D_PRJ-1_'+str(i)+'redrawn.jpg')"""
    
colorlist=['black','blue','crimson','darkorange','darkorchid','darkcyan','burlywood','deeppink','lawngreen','indigo','seashell','orangered','olive','powderblue','yellow']
    
linedict={}
circledict={}
for i in range(len(f)):
    for j in range(1,len(f[i])-1):
        for z in range(int(j+f[i][0]),len(f[i])-1+int(f[i][0])):
            if z in linedict.keys():
                linedict[z].append(((f[i][j],f[i][j+1]),colorlist[i%15]))
                circledict[z].append((f[i][j],colorlist[i%15]))
            else:
                linedict[z]=[((f[i][j],f[i][j+1]),colorlist[i%16])]
                circledict[z]=[(f[i][j],colorlist[i%16])]
                
    """for j in range(1,len(i)-1):
        draw.line((i[j],i[j+1]),fill='red')
    if i[0]==1:
        for j in range(1,len(i)-1):
            draw.line((i[j],i[j+1]),fill='blue')
    draw.text((i[1][0]-5,i[1][1]-5),str(f.index(i)),fill='white',font=font)"""
print circledict.keys(),linedict.keys()
for i in range(max(linedict.keys())):
    if i in linedict.keys():
        #c=names[i][:names[i].index('.')]+'_'+str(i)+'_raw.bmp'
        c=r'D:\ImageTesterOut\47A_03_R3D_D3D_PRJ-1_'+str(i)+'_TRF1.bmp'
        dd=Image.open(c)
        d=dd.convert('RGBA')
        #import ImageDraw,ImageFont
        drawee=ImageDraw.Draw(d)
        todraw=Image.new('RGBA',d.size,(0,0,0,0.5))
        draw=ImageDraw.Draw(todraw)
        for j in circledict[i]:
            #print j
            makecircle=[(j[0][0]+3,j[0][1]+0),(j[0][0]+3,j[0][1]+1),(j[0][0]+2,j[0][1]+2),(j[0][0]+1,j[0][1]+3),(j[0][0]+0,j[0][1]+3),(j[0][0]-1,j[0][1]+3),(j[0][0]-2,j[0][1]+2),(j[0][0]-3,j[0][1]+1),(j[0][0]-3,j[0][1]+0),(j[0][0]-3,j[0][1]-1),(j[0][0]-2,j[0][1]-2),(j[0][0]-1,j[0][1]-3),(j[0][0]-0,j[0][1]-3),(j[0][0]+1,j[0][1]-3),(j[0][0]+2,j[0][1]-2),(j[0][0]+3,j[0][1]-1)]
            draw.polygon(makecircle,outline=j[1])
        for k in linedict[i]:
            draw.line(j[0],fill=j[1],width=5)
        
        d.paste(todraw,(0,0),todraw)
        #print names[i][:names[i].index('.')]+'_'+str(i)+'movies.jpg'
        del draw        
        d.save(r'D:\20110714\Nuclei\47A\47A_03_R3D_D3D_PRJ-1_'+str(i)+'redrawn.jpg')
