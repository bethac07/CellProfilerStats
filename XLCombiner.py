"""Worksheet Combiner- part of CellProfiler Stats- by Beth Cimini Sept 2010

Combine the values of parameters listed excel spreadsheets into one large master file.
Note: In order for these to work, the first row (ie the column headings) cannot have any
duplicated values, including blank cells."""

import xlrd
import xlwt
from xlutils.copy import copy
import easygui as eg
from HandyXLModules import *



def addtomaster():
    a=eg.fileopenbox(msg='Choose Master Excel file',default='*.xls')
    reada=xlrd.open_workbook(a)
    writea=copy(reada)
    b=eg.fileopenbox(msg='Choose Excel file to be appended',default='*.xls')
    readb=xlrd.open_workbook(b)
    c=eg.enterbox('What identifier do you want to give the new file?')
    d=readsheets2files(reada,readb)
    t=[]
    for i in d[1]:
        if i in d[0]:
           g=colheadingreadernum(reada,d[0].index(i)) #read the column headers
           h=colheadingreadernum(readb,d[1].index(i)) #read the column headers
           if g==h: #If the column headers are identical...
                mastersheet=writea.get_sheet(d[0].index(i)) #pull the master sheet
                j=copysheet(reada,d[0].index(i)) #copy the sheet with that name from the first file
                k=copysheet(readb,d[1].index(i),1) #copy the sheet from the second file
                writesheet(mastersheet,k,len(j)) #copy the file just below the copy from the first file
                cc=[c]*len(k) #make the experiment identifier list
                writecol(mastersheet,cc,maxnumcols(k),len(j)) #add the experiment identifiers to each row
            
           else: #If the columns are NOT identical
                mastersheet=writea.get_sheet(d[0].index(i)) #pull the master sheet
                alen=len(copysheet(reada,d[0].index(i))) #find out how many rows are in the sheet from file 1
                blen=len(copysheet(readb,d[1].index(i))) #find out how many rows are in the sheet from file 2
                cc=[c]*(blen-1) #make the appropriate number of column identifiers
                m=g[:] #make a list of all the possible column headers
                for j in h:
                    if j not in m:
                        m.append(j)
                    else:
                        pass
                writecol(mastersheet,cc,len(m)-1,alen)
                for k in range(len(m)): # for each column header
                    if m[k] in h: #if it's in the second sheet, copy and write it beneath the first
                        hcol=copycol(readb,d[1].index(i),h.index(m[k]),1)
                        writecol(mastersheet,hcol,k,alen)
                    else:
                        pass
        else:
            t.append(i)
        
    if t!=[]: #if there are sheets which aren't in the master
        f=eg.multchoicebox(msg='Select which (if any) sheets to add to the master file',choices=t)
        if f!=[]: #if the user selects to add those sheets
            for i in f:
                mastersheet=writea.add_sheet(i)
                j=copysheet(readb,d[1].index(i))
                writesheet(mastersheet,j)
                cc=[c]*(len(j)-1)
                mastersheet.write(0,maxnumcols(j),'Experiment Identifier')
                writecol(mastersheet,cc,maxnumcols(j),1)
        else:
            pass
    writea.save(a)

def newmaster():
    a=eg.fileopenbox(msg='Choose 1st Excel file',default='*.xls')
    reada=xlrd.open_workbook(a)
    c=eg.enterbox('What identifier do you want to give this file?')
    b=eg.fileopenbox(msg='Choose 2nd Excel file',default='*.xls')
    readb=xlrd.open_workbook(b)
    d=eg.enterbox('What identifier do you want to give this file?')
    z=eg.filesavebox(msg='What do you want to name the master?',filetypes=["*.xls"])
    zxl=(z+'.xls') #automatically add the file extension
    master=xlwt.Workbook() #open a workbook for the master sheet
    e=readsheets2files(reada,readb) #read the sheet names
    z=e[0][:] #start a list with all of the sheet names of the first file
    if e[0]==e[1]: #If the names of the sheets are all identical, pass
        pass
    else:
        for i in e[1]: #If not, add any sheet names only in the second file to the earlier list
            if i not in e[0]:
                z.append(i)
    f=eg.multchoicebox('Which of these sheets do you want to use in the Master?',choices=z)
    for i in f: #For each sheet in the final sheet list
        if i in e[0] and colheadingreadernum(reada,e[0].index(i))!=[]: #If the sheet is in the first file and if it isn't blank
            g=colheadingreadernum(reada,e[0].index(i)) #read the column headers

            if i in e[1] and colheadingreadernum(readb,e[1].index(i))!=[]: #If the sheet is in the second file and it isn't blank
                h=colheadingreadernum(readb,e[1].index(i)) #read the column headers
                if g==h: #If the column headers are identical...
                    mastersheet=master.add_sheet(i) #add a sheet 
                    j=copysheet(reada,e[0].index(i)) #copy the sheet with that name from the first file
                    writesheet(mastersheet,j) #write it to the master sheet
                    cc=[c]*(len(j)-1) #make the experiment identifier a list as long as the number of entries for that experiment
                    mastersheet.write(0,maxnumcols(j),'Experiment Identifier') #write the heading
                    writecol(mastersheet,cc,maxnumcols(j),1) #add the experiment identifiers to each row
                    k=copysheet(readb,e[1].index(i),1) #copy the sheet from the second file
                    writesheet(mastersheet,k,len(j)) #copy the file just below the copy from the first file
                    dd=[d]*len(k) #make the experiment identifier list
                    writecol(mastersheet,dd,maxnumcols(k),len(j)) #add the experiment identifiers to each row
                else: #If the columns are NOT identical
                    mastersheet=master.add_sheet(i) #add a sheet
                    alen=len(copysheet(reada,e[0].index(i))) #find out how many rows are in the sheet from file 1
                    blen=len(copysheet(readb,e[1].index(i))) #find out how many rows are in the sheet from file 2
                    cc=[c]*(alen-1) #make the appropriate number of column identifiers
                    dd=[d]*(blen-1)
                    m=g[:] #make a list of all the possible column headers
                    for j in h:
                        if j not in m:
                            m.append(j)
                        else:
                            pass
                    writecol(mastersheet,cc,len(m),1) #add the experiment identifiers & heading
                    writecol(mastersheet,dd,len(m),alen)
                    mastersheet.write(0,len(m),'Experiment Identifier')
                    for k in range(len(m)): # for each column header
                        mastersheet.write(0,k,m[k]) #add the header
                        if m[k] in g: #if it's in the first sheet, copy and write it
                            gcol=copycol(reada,e[0].index(i),g.index(m[k]),1)
                            writecol(mastersheet,gcol,k,1)
                        else:
                            pass
                        if m[k] in h: #if it's in the second sheet, copy and write it beneath the first
                            hcol=copycol(readb,e[1].index(i),h.index(m[k]),1)
                            writecol(mastersheet,hcol,k,alen)
                        else:
                            pass
                   
            else: #If it's only in the first file, copy that
                mastersheet=master.add_sheet(i)
                j=copysheet(reada,e[0].index(i))
                writesheet(mastersheet,j)
                cc=[c]*(len(j)-1)
                mastersheet.write(0,maxnumcols(j),'Experiment Identifier')
                writecol(mastersheet,cc,maxnumcols(j),1)
        else:
            #If it's not in the first file, it must be in the second, so copy that
            mastersheet=master.add_sheet(i)
            j=copysheet(readb,e[1].index(i))
            writesheet(mastersheet,j)
            dd=[d]*(len(j)-1)
            mastersheet.write(0,maxnumcols(j),'Experiment Identifier')
            writecol(mastersheet,dd,maxnumcols(j),1)
    master.save(zxl)

def runxlcombiner():
    a=eg.boolbox('What do you want to do?',choices=['Add to an existing master','Create a new master'])
    if a:
        addtomaster()
    else:
        newmaster()
    
if __name__=='__main__':
    runxlcombiner()
