import xlwt
import os
from scipy import stats
import numpy
import easygui as eg
from HandyXLModules import *
import shelve

def choosewhichstats(book,default=False): #Choose which stats and set the filters
    pulldefs=shelve.open(os.path.join(os.curdir,'CompSinDoshelf'),writeback=True)
    if default==True:
        whichdefault=eg.choicebox('Which default set do you want to use?',choices=pulldefs.keys())
        g=pulldefs[whichdefault]
    else:
        a=readsheets1file(book) #Read the names of the sheets
        t=[]
        for i in a:
            if i != 'Image': #generally unhelpful output from cell profiler (has 150+ columns)- this could be removed or more sheets could be added
                ii=str('%.2d' %a.index(i)) #force the sheet index to be 2 digits- keeps everything in order
                c=colheadingreadernum(book,a.index(i)) #read the headings for each column
                for j in c: #for each sheet...
                    jj=str('%.3d' %c.index(j)) #force the column index to be 3 digits
                    t.append(ii+':)'+jj+':)'+str(i)+'-'+str(j)) #add each column to a master list
        if eg.ynbox(msg='Do you want to use the defaults?'):
            whichdefault=eg.choicebox('Which default set do you want to use?',choices=pulldefs.keys())
            g=pulldefs[whichdefault]
        else:
            f=eg.multchoicebox(msg='Select which parameters to analyze',choices=t)
            x=eg.multchoicebox(msg='Select which (if any) parameters to generate histograms for', choices=f)
            y=eg.multchoicebox(msg='Do you want to filter any of these?',choices=f)
            g=[]
            for i in f: #for all selected parameters
                s=[] #create a list for each parameter
                s.append(int(i[0:2])) #first item is the sheet index
                s.append(int(i[4:7])) #second item is the column index
                if i in x: #if the user wants a histogram, append 1, otherwise append 0
                    s.append(1)
                else:
                    s.append(0)
                if i in y: #if the user says they want to filter
                    numorstr=eg.boolbox(('How do you want to filter '+i[9:]),choices=['By numerical value','By experiment identifier'])
                    if numorstr: #if they say by numverical value
                        s.append(1) #create a numerical index for filtering type
                        filt=eg.multenterbox(fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))
                        s.append(filt) #add the type of filter to the list
                        unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+i[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                        if unfilt:
                            g.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                    else: 
                        s.append(2) #numerical index for the other filtering type
                        b=[]
                        c=colheadingreadernum(book,int(i[0:2])) #read the column headings for that sheet
                        d=copycol(book,int(i[0:2]),c.index('Experiment Identifier'),1) #find the column with the experiment identifier
                        for i in d:
                            if i not in b: #add only the unique identifiers to a list
                                b.append(i)
                        filt=eg.multchoicebox(msg='Which of these do you want to use in the analysis?',choices=b) #ask the users which of the identifiers they want to use
                        s.append(filt)
                        unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+i[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                        if unfilt:
                            g.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                else: #if the user does not want to filter, append 0
                    s.append(0)
                g.append(s) #append each annotated selection to the master list and return it
                if eg.ynbox('Do you want to save these settings as a new default?'):
                    newdefname=eg.enterbox(msg='Give this default a descriptive identifier')
                    pulldefs[newdefname]=g
        
    pulldefs.close()

    return g

def dothestuff(book,bookpath,writesheet,writesheetrows,default=False):
    c=choosewhichstats(book,default) #Choose the parameters to run statistics on
    writesheet.col(0).width=15000 #Make the first column wider to accomodate long parameter names
    row=writesheetrows+1 # set the starting row
    for i in c: #for all parameters
        statscol1=copycol(book,i[0],i[1]) #copy that column
        for m in range(len(statscol1)-1,0,-1):
            if statscol1[m]=='':
                del statscol1[m]
        if i[3]==1: #if the user chose to filter based on a numerical value
            statscol=[] # create an output list
            statscol.append(statscol1[0]) #add the column header
            statscol=statscol+sortfrominput(i[4],statscol1[1:]) #sort based on the user's input (i[4])
            i[4]=i[4][0]+i[4][1]
        elif i[3]==2: #if the user chose to filter based on an experiment identifier
            headings=colheadingreadernum(book,i[0]) #look at the headers
            identcol=copycol(book,i[0],headings.index('Experiment Identifier')) #find the column for the experiment identifier
            okrows=[] #start an empty list
            for j in range(len(identcol)): #copy all of the row numbers that meet the users specifications (i[4])
                if identcol[j] in i[4]:
                    okrows.append(j)
            statscol=[] #start a new list
            statscol.append(statscol1[0]) #add the column header
            for k in range(1,len(statscol1)): #for all of the values
                if k in okrows: #if the value comes from a selected row, add it to the list
                    statscol.append(statscol1[k])
        else:
            statscol=statscol1 #if the user didn't choose to filter, just copy the list verbatim
        statscolheading= statscol[0] #pull the heading for that parameter
        writesheet.write(row,1,'mean') #set up column headings
        writesheet.write(row,2,'+/-')
        writesheet.write(row,3,'st dev')
        writesheet.write(row,4,'median')
        writesheet.write(row+1,0,'n='+str((len(statscol)-1))) #write the number of values in the dataset
        mean=numpy.mean(statscol[1:]) #calculate and write the mean...
        writesheet.write(row+1,1,mean)
        stdev=numpy.std(statscol[1:]) #... standard deviation...
        writesheet.write(row+1,2,'+/-')
        writesheet.write(row+1,3,stdev)
        median=numpy.median(statscol[1:]) #and median
        writesheet.write(row+1,4,median)
        a=readsheets1file(book) #Read the names of the sheets
        if len(i)==5: #if the user chose a filter, write what it was
            writesheet.write(row,0,a[i[0]]+'-'+statscolheading+','+str(i[4])) #write the name of the parameter we're looking at
            histtitle=(a[i[0]]+'-'+statscolheading+','+str(i[4]))
        else:
            histtitle=(a[i[0]]+'-'+statscolheading)
            writesheet.write(row,0,a[i[0]]+'-'+statscolheading)
        if i[2]==1: #if the user chose to use a histogram
            writesheet.row(row).set_style(xlwt.easyxf('font:height 5000')) #make the row tall enough to accomodate it
            saveas=bookpath+statscolheading
            graphhist(statscol,histtitle,saveas,statscolheading)
            writesheet.insert_bitmap(saveas+'.bmp',row,6) #write the .bmp to the Excel sheet
            os.remove(saveas+'.png')
            os.remove(saveas+'.bmp')
        row=row+3 #move 3 rows down for the next parameter
    return row
