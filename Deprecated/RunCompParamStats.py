"""Run Comparative Parameter Statistics- part of CellProfiler Stats- by Beth Cimini Sept 2010

Run statistics based on Excel sheet inputs- currently set for n, mean, standard deviation and median but others could be added as well."""

import xlrd
import xlwt
from xlutils.copy import copy
import os
import easygui as eg
from HandyXLModules import *
import CompParamXY as CPXY
import CompParamSingle as CPS

def singleorcomp(book,book2,ident,ident2,bookpath,writesheet,writesheetrows,filter1):
    xvxy=eg.boolbox('Do you want to compare 2 single parameters, or 2 relative parameters (ie y/x)?', choices=('single','Y/X'))
    if xvxy:
        a=CPS.dothestuffsin(book,book2,ident,ident2,bookpath,writesheet,writesheetrows,filter1)
    else:
        a=CPXY.dothestuffxy(book,book2,ident,ident2,bookpath,writesheet,writesheetrows,filter1)
    cont=eg.ynbox(msg='Do other comparisons with same input/output files?')
    if cont:
        singleorcomp(book,book2,ident,ident2,bookpath,writesheet,a,filter1)
    else:
        pass

def runCompParFileIO():
    howmany=eg.boolbox(msg='Are the parameters in the same or different files?',choices=('Same','Different'))
    aa=eg.fileopenbox(msg='Choose input Excel file',default='*.xls') #select a file, open it
    ident=eg.enterbox('What should the identifier for the first data set be?')
    if howmany:
        bb=aa
        filter1=1
    else:
        bb=eg.fileopenbox(msg='Choose input Excel file',default='*.xls') #select a file, open it
        filter1=0
    ident2=eg.enterbox('What should the identifier for the second data set be?')
    book2=xlrd.open_workbook(bb)
    book=xlrd.open_workbook(aa)
    bookpath=os.path.dirname(aa)+'/'
    sheetnames=readsheets1file(book) #read the sheet names
    writesheetrows=0
    reuse=eg.indexbox('Where to save the output?',choices=['Add to a new sheet in the (first) input file',
                                                           'Add to an existing sheet in the (first) input file','Create a new file', 
                                                           'Add to a new sheet in another (or the second) file','Add to an existing sheet in another (or the second) file'])
    if reuse==0: #If the user chooses to add a new sheet to the existing file
        writebook=copy(book) #copy to xlutils to make wrtable
        newsheetname=eg.enterbox('What do you want to call the new sheet?')
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
    singleorcomp(book,book2,ident,ident2,bookpath,writesheet,writesheetrows,filter1)
    if reuse==0 or reuse==1: #if we're saving the input file, save it with the input name
        writebook.save(aa)
    elif reuse==2: #if we're saving a new file, save it with the name the user input
        writebook.save(w)
    elif reuse==3 or reuse==4: #if we're saving an old file, save it with it's original name
        writebook.save(o)

if __name__=='__main__':
    runCompParFileIO()  
