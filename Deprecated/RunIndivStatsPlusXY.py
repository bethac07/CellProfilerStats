"""Run Individual Statistics- part of CellProfiler Stats- by Beth Cimini Sept 2010

Run statistics based on Excel sheet inputs- currently set for n, mean, standard deviation and median but others could be added as well."""

import xlrd
import xlwt
from xlutils.copy import copy
import os
import easygui as eg
from HandyXLModules import *
import CompXYDo as CXYD
import CompSinDo as CSD

def singleorcomp(book,bookpath,writesheet,writesheetrows):
    xvxy=eg.boolbox('Do you want to do statistics on single parameters or relative parameters (ie y/x)?', choices=('single','Y/X'))
    if xvxy:
        a=CSD.dothestuff(book,bookpath,writesheet,writesheetrows)
    else:
        a=CXYD.dothestuff(book,bookpath,writesheet,writesheetrows)
    cont=eg.ynbox(msg='Do other statistics with same input/output files?')
    if cont:
        singleorcomp(book,bookpath,writesheet,a)
    else:
        pass

def runindivstats(aa=0,defuse='a',name='blank',sinorcomp=0,default=False):
    if aa==0:
        aa=eg.fileopenbox(msg='Choose input Excel file',default='*.xls') #select a file, open it
    book=xlrd.open_workbook(aa)
    bookpath=os.path.dirname(aa)+'/'
    writesheetrows=0 #set initial value of which row to write in as 0- is modified below if the user chooses to add to an existing sheet
    if defuse=='a':
        reuse=eg.indexbox('Where to save the output?',choices=['Add to a new sheet in the input file',
                                                           'Add to an existing sheet in the input file','Create a new file', 
                                                           'Add to a new sheet in another file','Add to an existing sheet in another file'])
    else:
        reuse=0
    if reuse==0: #If the user chooses to add a new sheet to the existing file
        writebook=copy(book) #copy to xlutils to make wrtable
        if name=='blank':
            newsheetname=eg.enterbox('What do you want to call the new sheet?')
        else:
            newsheetname='Statistics'
        writesheet=writebook.add_sheet(newsheetname) #add the new sheet
        writesheet.write(0,0,'Statistics') #put "Statistics" as a header
    if reuse==1: #If the user chooses to add to an exiting sheet
        writebook=copy(book) #make the file writable
        whichsheet=sheetnames.index(eg.choicebox(msg='Which sheet?',choices=sheetnames)) #select which sheet to add to
        writesheet=writebook.get_sheet(whichsheet) #create a writable version of the sheet
        readsheet=book.sheet_by_index(whichsheet) #create a readable version of the sheet to find the number of rows
        writesheetrows=(readsheet.nrows) #set a new starting row based on how many rows already present
    if reuse==2: #If the user wants to make a new file
        w=eg.filesavebox(msg='What do you want to name the file?',filetypes=["*.xls"])+'.xls'
        writebook=xlwt.Workbook() #make a new file
        newsheetname=eg.enterbox('What do you want to call the new sheet?')
        writesheet=writebook.add_sheet(newsheetname) #give it a sheet
        writesheet.write(0,0,'Statistics') #put "Statistics" as a header
    if reuse==3: #If the user wants a new sheet in an old file
        o=eg.fileopenbox(msg='Choose Excel file to write to',default='*.xls')
        otherbook=xlrd.open_workbook(o) #Open it
        writebook=copy(otherbook) #make it writable
        newsheetname=eg.enterbox('What do you want to call the new sheet?')
        writesheet=writebook.add_sheet(newsheetname) #add the new sheet
        writesheet.write(0,0,'Statistics') #put "Statistics as a header"
    if reuse==4: #If the user wants to write to an existing sheet in an existing file
        o=eg.fileopenbox(msg='Choose Excel file to write to',default='*.xls')
        otherbook=xlrd.open_workbook(o) #Open it
        writebook=copy(otherbook)#Make it writable
        othersheetnames=readsheets1file(otherbook) #find out the names of the sheets
        whichsheet=othersheetnames.index(eg.choicebox(msg='Which sheet?',choices=othersheetnames)) #choose which sheets
        writesheet=writebook.get_sheet(whichsheet) #create a writable version of the sheet
        readsheet=otherbook.sheet_by_index(whichsheet) #create a readable version to find the number of rows
        writesheetrows=(readsheet.nrows) #set a new starting row based on how many rows are already present.
    if sinorcomp==0:
        singleorcomp(book,bookpath,writesheet,writesheetrows)
    else:
        if default==False:
            a=CSD.dothestuff(book,bookpath,writesheet,writesheetrows)
        else:
            a=CSD.dothestuff(book,bookpath,writesheet,writesheetrows,default=True)
    if reuse==0 or reuse==1: #if we're saving the input file, save it with the input name
        writebook.save(aa)
    elif reuse==2: #if we're saving a new file, save it with the name the user input
        writebook.save(w)
    elif reuse==3 or reuse==4: #if we're saving an old file, save it with it's original name
        writebook.save(o)

def batchsincomps(direct):
    a=direct
    for i in os.listdir(a):
        if '.xls' not in i:
            for j in os.listdir(os.path.join(a,i)):
                if '.xls' in j:
                    runindivstats(os.path.join(a,i,j),defuse=0,name='Statistics',sinorcomp=1,default=True)

if __name__=='__main__':
    runindivstats()
