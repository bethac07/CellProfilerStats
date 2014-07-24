import xlwt
import os
import random
import numpy
from scipy import stats
import easygui as eg
from HandyXLModules import *
import shelve


map2={}
t=[]
colsbysheet=[]

def choosewhichstatssin(book,book2,filter1):
    pulldefs=shelve.open(os.path.join(os.curdir,'CompParamSingleshelf'),writeback=True)
    if eg.ynbox(msg='Do you want to use the defaults?'):
        whichdefault=eg.choicebox('Which default set do you want to use?',choices=pulldefs.keys())
        g=pulldefs[whichdefault]
    else:
    #Pull all the possible parameters
        if map2=={}:
            global t
            global colsbysheet
            a=readsheets1file(book)
            colsbysheet=[] #create a list of parameters sorted by sheet- for choosing the y value once x is chosen
            t=[] #create a list of all the parameters
            b2=readsheets1file(book2)
            t2=[]
            t2stripped=[]
    
            for m in b2:
                mm=colheadingreadernum(book2,b2.index(m))
                for n in mm:
                    t2stripped.append(m+n)
                    t2.append((b2.index(m),mm.index(n)))
            for i in a:
                if i != 'Image': #generally unhelpful output from cell profiler (has 150+ columns)- this could be removed or more sheets could be added
                    ii=str('%.2d' %a.index(i))
                    c=colheadingreadernum(book,a.index(i))
                    newsheet=[]
                    for j in c:
                        if i+j in t2stripped:
                            map2[(a.index(i),c.index(j))]=t2[t2stripped.index(i+j)]
                            jj=str('%.3d' %c.index(j))
                            t.append(ii+':)'+jj+':)'+str(i)+'-'+str(j))
                            newsheet.append(ii+':)'+jj+':)'+str(i)+'-'+str(j))
                    colsbysheet.append(newsheet)
        if filter1==1:
            b=[]
            c=colheadingreadernum(book,int(fx[0:2])) #read the column headings for that sheet
            c2=colheadingreadernum(book2,map2[(int(fx[0:2]),int(fx[4:7]))][0])
            d=copycol(book,int(fx[0:2]),c.index('Experiment Identifier'),1) #find the column with the experiment identifier
            d2=copycol(book2,map2[(int(fx[0:2]),int(fx[4:7]))][0],c2.index('Experiment Identifier'),1)
            for ii in d:
                if ii in d2:
                    if ii not in b: #add only the unique identifiers to a list
                        b.append(ii)
            filt1=eg.multchoicebox(msg='Which of these do you want to use for condition 1?',choices=b)#ask the users which of the identifiers they want to use
            filt2=eg.multchoicebox(msg='Which of these do you want to use for condition 2?',choices=b)
       #Handle the x parameter
        
    
        else:
            xparams=[]
            fx=eg.multchoicebox(msg='Select parameters',choices=t)
            graphx=eg.multchoicebox(msg='Select which (if any) parameters to graph', choices=fx)
            if filter1==1:
                filtf=eg.multchoicebox(msg='Do you want also want to numerically filter any of these?',choices=(fx))
            else:
                filtf=eg.multchoicebox(msg='Do you want to filter any of these?',choices=(fx))
            for param in xrange(len(fx)):
                s=[]
                s.append(int(fx[param][0:2])) #first item is the sheet index
                s.append(int(fx[param][4:7])) #second item is the column index
                if fx[param] in graphx:
                    s.append(1)#pass the graphing option
                else:
                    s.append(0)
                if filter1==0:
                    if fx[param] in filtf: #if the user wants to filter
                        numorstr=eg.boolbox(('How do you want to filter '+fx[param][9:]),choices=['By numerical value','By experiment identifier'])
                        if numorstr: #if they say by numverical value
                            s.append(1) #create a numerical index for filtering type
                            filt=eg.multenterbox(fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))#let the user input the filter they want
                            s.append(filt) #add the type of filter to the list
                            unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[param][9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                            if unfilt:
                                xparams.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                        else: 
                            s.append(2) #numerical index for the other filtering type
                            b=[]
                            c=colheadingreadernum(book,int(fx[param][0:2])) #read the column headings for that sheet
                            c2=colheadingreadernum(book2,map2[(int(fx[param][0:2]),int(fx[param][4:7]))][0])
                            d=copycol(book,int(fx[param][0:2]),c.index('Experiment Identifier'),1) #find the column with the experiment identifier
                            d2=copycol(book2,map2[(int(fx[param][0:2]),int(fx[param][4:7]))][0],c2.index('Experiment Identifier'),1)
                            for ii in d:
                                if ii in d2:
                                    if ii not in b: #add only the unique identifiers to a list
                                        b.append(ii)
                            filt=eg.multchoicebox(msg='Which of these do you want to use in the analysis?',choices=b) #ask the users which of the identifiers they want to use
                            s.append(filt)
                            unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[param]+'?', title=' ', choices=('Yes', 'No'), image=None)
                            if unfilt:
                                xparams.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                    else: #if the user does not want to filter, append 0
                        s.append(0)
                    
                if filter1==1:
                    if fx[param] in filtf:
                        s.append(3) #create a numerical index for filtering type
                        filt=eg.multenterbox(fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))#let the user input the filter they want
                        s.append((filt1,filt)) #add the type of filter to the list
                        unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[param][9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                        if unfilt:
                            xparams.append((s[0:3]+[2,filt1])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                    else:
                        s=s+[2,filt1]
                xparams.append(s)
            
            xparams2=[]
            for i in xrange(len(xparams)):
                mapped=map2[(xparams[i][0],xparams[i][1])]
                if filter1==1:
                    if xparams[i][3]==2:
                        xparams2.append([mapped[0],mapped[1],xparams[i][2],2,filt2])
                    if xparams[i][3]==3:
                        xparams2.append([mapped[0],mapped[1],xparams[i][2],3,(filt2,xparams[i][4][1])])
                else:
                    xparams2.append([mapped[0],mapped[1]]+xparams[i][2:])
                
            g=[xparams,xparams2] #return the list of parameters the user wants to compare
            if eg.ynbox('Do you want to save these settings as a new default?'):
                newdefname=eg.enterbox(msg='Give this default a descriptive identifier')
                pulldefs[newdefname]=g
        
    pulldefs.close()
    return g

def arrange2sin(book,book2,filter1):
    whichparams=choosewhichstatssin(book,book2,filter1)
    t=[]
    first=arrangedivsin(book,whichparams[0])
    t.append(first)
    second=arrangedivsin(book2,whichparams[1])
    t.append(second)
    return t

def dothestuffsin(book,book2,ident,ident2,bookpath,writesheet,writesheetrows,filter1):
    z=arrange2sin(book,book2,filter1)
    c=z[0] #Pull the parameters to run statistics on
    d=z[1]
    writesheet.col(0).width=15000 #Make the first column wider to accomodate long parameter names
    row=writesheetrows+1 # set the starting row
    for i in xrange(len(c)):
        title=c[i][0][0]
        writesheet.write(row,0,title)
        writesheet.write(row+1,0,ident)
        writesheet.write(row+2,0,'n='+str(len(c[i])-1)) #write the calculated parameters, including n...
        writesheet.write(row+1,1,'mean') #set up column headings
        writesheet.write(row+1,2,'+/-')
        writesheet.write(row+1,3,'st dev')
        writesheet.write(row+1,4,'median')
        mean=numpy.mean(c[i][1:]) #calculate and write the mean...
        writesheet.write(row+2,1,mean)
        stdev=numpy.std(c[i][1:]) #... standard deviation...
        writesheet.write(row+2,2,'+/-')
        writesheet.write(row+2,3,stdev)
        median=numpy.median(c[i][1:]) #and median
        writesheet.write(row+2,4,median)
        uvalue,pvalue=stats.mannwhitneyu(c[i][1:],d[i][1:])
        writesheet.write(row+1,6,'u value')
        writesheet.write(row+2,6,uvalue)
        writesheet.write(row+1,7,'p value')
        writesheet.write(row+2,7,pvalue*2)
        if c[i][0][1]==1: #if the user chose to graph the results
            rint=random.randint(0,999999)
            writesheet.row(row).set_style(xlwt.easyxf('font:height 5000')) #make the row tall enough to accomodate it
            graph2hists(c[i][1:],d[i][1:],ident,ident2,title,bookpath+str(rint),uvalue,pvalue*2,title) #graph it---CUSTOMIZABLE HERE
            writesheet.insert_bitmap(bookpath+str(rint)+'.bmp',row,6) #Put the figure into the excel sheet
            os.remove(bookpath+str(rint)+'.png')
            os.remove(bookpath+str(rint)+'.bmp')
        row=row+3 #move down 4 rows for the next parameter
        writesheet.write(row,0,ident2)
        writesheet.write(row+1,0,'n='+str(len(d[i])-1)) #write the calculated parameters, including n...
        writesheet.write(row,1,'mean') #set up column headings
        writesheet.write(row,2,'+/-')
        writesheet.write(row,3,'st dev')
        mean=numpy.mean(d[i][1:]) #calculate and write the mean...
        writesheet.write(row+1,1,mean)
        stdev=numpy.std(d[i][1:]) #... standard deviation...
        writesheet.write(row+1,2,'+/-')
        writesheet.write(row+1,3,stdev)
        median=numpy.median(d[i][1:]) #and median
        writesheet.write(row+1,4,median)
        row=row+3 #move down 3 rows for the next parameter
    return row
