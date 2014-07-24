from HandyXLModules import *
import xlwt
import os
import random
import numpy
from scipy import stats
import easygui as eg
import shelve

map2={}
t=[]
colsbysheet=[]

def choosewhichstatsxy(book,book2,filter1):
    pulldefs=shelve.open(os.path.join(os.curdir,'CompParamXYshelf'),writeback=True)
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
        #Handle the x parameter
        xparams=[]
        fx=eg.choicebox(msg='Select an x-axis parameter',choices=t)
        s=[]
        s.append(int(fx[0:2])) #first item is the sheet index
        s.append(int(fx[4:7])) #second item is the column index
        s.append(1)#placeholder for the graphing option
        if filter1==0:
            filtf=eg.ynbox(msg='Do you want to filter the x axis?')
            if filtf: #if the user wants to filter
                numorstr=eg.boolbox(('How do you want to filter '+fx[9:]),choices=['By numerical value','By experiment identifier'])
                if numorstr: #if they say by numverical value
                    s.append(1) #create a numerical index for filtering type
                    filt=eg.multenterbox(fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))#let the user input the filter they want
                    s.append(filt) #add the type of filter to the list
                    unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                    if unfilt:
                        xparams.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                else: 
                    s.append(2) #numerical index for the other filtering type
                    b=[]
                    c=colheadingreadernum(book,int(fx[0:2])) #read the column headings for that sheet
                    c2=colheadingreadernum(book2,map2[(int(fx[0:2]),int(fx[4:7]))][0])
                    d=copycol(book,int(fx[0:2]),c.index('Experiment Identifier'),1) #find the column with the experiment identifier
                    d2=copycol(book2,map2[(int(fx[0:2]),int(fx[4:7]))][0],c2.index('Experiment Identifier'),1)
                    for ii in d:
                        if ii in d2:
                            if ii not in b: #add only the unique identifiers to a list
                                b.append(ii)
                    filt=eg.multchoicebox(msg='Which of these do you want to use in the analysis?',choices=b) #ask the users which of the identifiers they want to use
                    s.append(filt)
                    unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                    if unfilt:
                        xparams.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
            else: #if the user does not want to filter, append 0
                s.append(0)
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
            filt3=eg.ynbox(msg='Do you also want to use a numeric filter?')
            if filt3:
                s.append(3) #create a numerical index for filtering type
                filt=eg.multenterbox(fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))#let the user input the filter they want
                s.append((filt1,filt)) #add the type of filter to the list
                unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
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
                    xparams2.append([mapped[0],mapped[1],1,2,filt2])
                if xparams[i][3]==3:
                    xparams2.append([mapped[0],mapped[1],1,3,(filt2,xparams[i][4][1])])
            else:
                xparams2.append([mapped[0],mapped[1]]+xparams[i][2:])
            
        #Handle the y parameter
        for i in range(len(colsbysheet)):
            if colsbysheet[i][0][0:2]==fx[0:2]:
                colsbysheetindex=i
        fy=eg.multchoicebox(msg='Select 1 or more y-axis parameters (run individually)',choices=colsbysheet[colsbysheetindex]) #show all the parameters from the sheet with the x parameter
        graphy=eg.multchoicebox(msg='Select which (if any) parameters to graph', choices=fy)
        y=eg.multchoicebox(msg='Do you want to filter any of these?',choices=(fy))
        yparams=[]
        for i in fy:
            s=[]
            s.append(int(i[0:2])) #sheet index
            s.append(int(i[4:7])) #column index
            if i in graphy: #mark whether the user wants to graph the output or not
                s.append(1)
            else:
                s.append(0)
            if i in y:
                numorstr=eg.boolbox(('How do you want to filter '+i[9:]),choices=['By numerical value','By experiment identifier']) #same filtering algorithm as above
                if numorstr:
                    s.append(1)
                    filt=eg.multenterbox(fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))
                    s.append(filt)
                    unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+i[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                    if unfilt:
                        yparams.append((s[0:3]+[0]))
                else:
                    s.append(2)
                    b=[]
                    c=colheadingreadernum(book,int(i[0:2]))
                    d=copycol(book,int(i[0:2]),c.index('Experiment Identifier'),1)
                    for i in d:
                        if i not in b:
                            b.append(i)
                    filt=eg.multchoicebox(msg='Which of these do you want to use in the analysis?',choices=b)
                    s.append(filt)
                    unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+i[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                    if unfilt:
                        yparams.append((s[0:3]+[0]))
            else:
                s.append(0)
    
            yparams.append(s)
        yparams2=[]
        for i in xrange(len(yparams)):
            mapped=map2[(yparams[i][0],yparams[i][1])]
            yparams2.append([mapped[0],mapped[1]]+yparams[i][2:])
        g=[xparams,yparams,xparams2,yparams2] #return the list of which x and y parameters the user wants to compare
        if eg.ynbox('Do you want to save these settings as a new default?'):
            newdefname=eg.enterbox(msg='Give this default a descriptive identifier')
            pulldefs[newdefname]=g
        
    pulldefs.close()
    return g

def arrangedivxy(book,a0,a1):
    xaxes=a0 #separate x from y
    yaxes=a1
    statstorun=[]
    sh=readsheets1file(book)
    for x in xaxes: #for x (or filtered and unfiltered if the user so chose)
        xvalstart=[]
        xvalstart=copycol(book,x[0],x[1]) #copy that column
        if x[3]==1: #if the user chose to filter based on a numerical value
            statsx=[] # create an output list
            statsx.append(sh[x[0]]+'-'+xvalstart[0]+'('+x[4][0]+x[4][1]+')') #add the column header and filter
            statsx=statsx+conservsortfrominput(x[4],xvalstart[1:]) #sort based on the user's input (i[4])- keeps the indices the same by appending '' if the number is filtered out
        elif x[3]==2: #if the user chose to filter based on an experiment identifier
            headings=colheadingreadernum(book,x[0]) #look at the headers
            identcol=copycol(book,x[0],headings.index('Experiment Identifier')) #find the column for the experiment identifier
            okrows=[] #start an empty list
            for j in xrange(len(identcol)): #copy all of the row numbers that meet the users specifications (i[4])
                if identcol[j] in x[4]:
                    okrows.append(j)
            statsx=[] #start a new list
            statsx.append(sh[x[0]]+'-'+xvalstart[0]+'('+str(x[4])+')') #add the column header and filter
            for k in xrange(1,len(xvalstart)): #for all of the values
                if k in okrows: #if the value comes from a selected row, add it to the list
                    statsx.append(xvalstart[k])
                else: #otherwise, append a placeholder so that you can directly compare the indices to the list of y values
                    statsx.append('')
        elif x[3]==3: #if the user chose to filter based on an experiment identifier AND a number
                headings=colheadingreadernum(book,x[0]) #look at the headers
                identcol=copycol(book,x[0],headings.index('Experiment Identifier')) #find the column for the experiment identifier
                okrows=[] #start an empty list
                for j in xrange(len(identcol)): #copy all of the row numbers that meet the users specifications (i[4])
                    if identcol[j] in x[4][0]:
                        okrows.append(j)
                statsint=[] #start a new list
                statsint.append(sh[x[0]]+'-'+xvalstart[0]+'('+str(x[4][0])+'+'+str(x[4][1][0])+str(x[4][1][1])+')') #add the column header and filter
                for k in xrange(1,len(xvalstart)): #for all of the values
                    if k in okrows: #if the value comes from a selected row, add it to the list
                        statsint.append(xvalstart[k])
                    else:
                        statsint.append('') #conserve the index
                statsx=[] # create an output list
                statsx.append(statsint[0]) #add the column header and filter
                statsx=statsx+conservsortfrominput(x[4][1],statsint[1:]) #sort based on the user's input (i[4]), conserving the index
        else:
            statsx=[sh[x[0]]+'-'+xvalstart[0]]
            statsx=statsx+xvalstart[1:] #if the user didn't choose to filter, just copy the list verbatim
        for y in yaxes: #same logic as above, but for y
            yvalstart=[]
            yvalstart=copycol(book,y[0],y[1]) #copy that column
            if y[3]==1: #if the user chose to filter based on a numerical value
                statsy=[] # create an output list
                statsy.append(sh[y[0]]+'-'+yvalstart[0]+'('+y[4][0]+y[4][1]+')') #add the column header and filter
                statsy=statsy+conservsortfrominput(y[4],yvalstart[1:]) #sort based on the user's input (i[4]), conserving the index
            elif y[3]==2: #if the user chose to filter based on an experiment identifier
                headings=colheadingreadernum(book,y[0]) #look at the headers
                identcol=copycol(book,y[0],headings.index('Experiment Identifier')) #find the column for the experiment identifier
                okrows=[] #start an empty list
                for j in xrange(len(identcol)): #copy all of the row numbers that meet the users specifications (i[4])
                    if identcol[j] in y[4]:
                        okrows.append(j)
                statsy=[] #start a new list
                statsy.append(sh[y[0]]+'-'+yvalstart[0]+'('+str(y[4])+')') #add the column header and filter
                for k in xrange(1,len(yvalstart)): #for all of the values
                    if k in okrows: #if the value comes from a selected row, add it to the list
                        statsy.append(yvalstart[k])
                    else:
                        statsy.append('') #conserve the index
            else:
                statsy=[sh[y[0]]+'-'+yvalstart[0]]
                statsy=statsy+yvalstart[1:] #if the user didn't choose to filter, just copy the list verbatim
            divfinal=[[statsx[0],statsy[0],y[2]]]
            for i in xrange(1,len(statsx)): #If BOTH the x and y parameters were not filtered out (ie still numbers), add to the final list
                if type(statsx[i])==float:
                    if type(statsy[i])==float:
                        divfinal.append(statsy[i]/statsx[i])
                    else:
                        pass
                else:
                    pass
            statstorun.append(divfinal) #append the final list for each parameter to a master list
    return statstorun

def arrange2divxy(book,book2,filter1):
    whichparams=choosewhichstatsxy(book,book2,filter1)
    t=[]
    first=arrangedivxy(book,whichparams[0],whichparams[1])
    t.append(first)
    second=arrangedivxy(book2,whichparams[2],whichparams[3])
    t.append(second)
    return t

def dothestuffxy(book,book2,ident,ident2,bookpath,writesheet,writesheetrows,filter1):
    z=arrange2divxy(book,book2,filter1)
    c=z[0] #Pull the parameters to run statistics on
    d=z[1]
    writesheet.col(0).width=15000 #Make the first column wider to accomodate long parameter names
    row=writesheetrows+1 # set the starting row
    for i in xrange(len(c)):
        xtitlehalf=c[i][0][0]+' vs.'
        writesheet.write(row,0,'X='+c[i][0][0])
        ytitlehalf=c[i][0][1]
        writesheet.write(row+1,0,'Y='+ytitlehalf)
        writesheet.write(row+2,0,ident)
        writesheet.write(row+3,0,'n='+str(len(c[i])-1)) #write the calculated parameters, including n...
        writesheet.write(row+2,1,'mean') #set up column headings
        writesheet.write(row+2,2,'+/-')
        writesheet.write(row+2,3,'st dev')
        writesheet.write(row+2,4,'median')
        mean=numpy.mean(c[i][1:]) #calculate and write the mean...
        writesheet.write(row+3,1,mean)
        stdev=numpy.std(c[i][1:]) #... standard deviation...
        writesheet.write(row+3,2,'+/-')
        writesheet.write(row+3,3,stdev)
        median=numpy.median(c[i][1:]) #and median
        writesheet.write(row+3,4,median)
        uvalue,pvalue=stats.mannwhitneyu(c[i][1:],d[i][1:])
        writesheet.write(row+1,6,'u value')
        writesheet.write(row+2,6,uvalue)
        writesheet.write(row+1,7,'p value')
        writesheet.write(row+2,7,pvalue*2)
        if c[i][0][2]==1: #if the user chose to graph the results
            title=xtitlehalf+ytitlehalf #create a title
            rint=random.randint(0,999999)
            writesheet.row(row).set_style(xlwt.easyxf('font:height 5000')) #make the row tall enough to accomodate it
            graph2hists(c[i][1:],d[i][1:],ident,ident2,title,bookpath+str(rint),uvalue,pvalue*2,xtitlehalf+ytitlehalf) #graph it---CUSTOMIZABLE HERE
            writesheet.insert_bitmap(bookpath+str(rint)+'.bmp',row,6) #Put the figure into the excel sheet
            os.remove(bookpath+str(rint)+'.png')
            os.remove(bookpath+str(rint)+'.bmp')
        row=row+4 #move down 4 rows for the next parameter
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








