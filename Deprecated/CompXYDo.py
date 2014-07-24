import xlwt
import numpy as np
import os
from scipy import stats
import easygui as eg
from HandyXLModules import *
import shelve

t=[]
colsbysheet=[]
def choosewhichstatsxy(book):
    pulldefs=shelve.open(os.path.join(os.curdir,'CompXYDoshelf'),writeback=True)
    if eg.ynbox(msg='Do you want to use the defaults?'):
        whichdefault=eg.choicebox('Which default set do you want to use?',choices=pulldefs.keys())
        g=pulldefs[whichdefault]
    else:
        #Pull all the possible parameters
        global colsbysheet
        if colsbysheet==[]:
            a=readsheets1file(book)
            global t
            colsbysheet=[] #create a list of parameters sorted by sheet- for choosing the y value once x is chosen
            t=[] #create a list of all the parameters
            for i in a:
                if i != 'Image': #generally unhelpful output from cell profiler (has 150+ columns)- this could be removed or more sheets could be added
                    ii=str('%.2d' %a.index(i))
                    c=colheadingreadernum(book,a.index(i))
                    newsheet=[]
                    for j in c:
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
                d=copycol(book,int(fx[0:2]),c.index('Experiment Identifier'),1) #find the column with the experiment identifier
                for ii in d:
                    if ii not in b: #add only the unique identifiers to a list
                        b.append(ii)
                filt=eg.multchoicebox(msg='Which of these do you want to use in the analysis?',choices=b) #ask the users which of the identifiers they want to use
                unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+fx[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                if unfilt:
                    xparams.append((s[0:3]+[0])) #if the user wants an unfiltered version, add the sheet,column,and histogram preference with the index for unfiltered (0)
                s.append(filt)
        else: #if the user does not want to filter, append 0
            s.append(0)
        xparams.append(s)
    
        #Handle the y parameter
        for i in range(len(colsbysheet)):
            if colsbysheet[i][0][0:2]==fx[0:2]:
                colsbysheetindex=i
        fy=eg.multchoicebox(msg='Select 1 or more y-axis parameters (run individually)',choices=colsbysheet[colsbysheetindex])#show all the parameters from the sheet with the x parameter
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
                    unfilt=eg.ynbox(msg='Do you also want to add an unfiltered version of '+i[9:]+'?', title=' ', choices=('Yes', 'No'), image=None)
                    if unfilt:
                        yparams.append((s[0:3]+[0]))
                    s.append(filt)
            else:
                s.append(0)
    
            yparams.append(s)
        g=[xparams,yparams] #return the list of which x and y parameters the user wants to compare
        if eg.ynbox('Do you want to save these settings as a new default?'):
            newdefname=eg.enterbox(msg='Give this default a descriptive identifier')
            pulldefs[newdefname]=g
        
    pulldefs.close()
    return g

def arrangexy(book):
    listxy=choosewhichstatsxy(book) #get the list of x and y parameters
    xaxes=listxy[0] #separate x from y
    yaxes=listxy[1]
    statstorun=[]
    a=readsheets1file(book)
    for x in xaxes: #for x (or filtered and unfiltered if the user so chose)
        xvalstart=[]
        xvalstart=copycol(book,x[0],x[1]) #copy that column
        if x[3]==1: #if the user chose to filter based on a numerical value
            statsx=[] # create an output list
            statsx.append(a[x[0]]+'-'+xvalstart[0]+'('+x[4][0]+x[4][1]+')') #add the column header and filter
            statsx=statsx+conservsortfrominput(x[4],xvalstart[1:]) #sort based on the user's input (i[4])- keeps the indices the same by appending '' if the number is filtered out
        elif x[3]==2: #if the user chose to filter based on an experiment identifier
            headings=colheadingreadernum(book,x[0]) #look at the headers
            identcol=copycol(book,x[0],headings.index('Experiment Identifier')) #find the column for the experiment identifier
            okrows=[] #start an empty list
            for j in range(len(identcol)): #copy all of the row numbers that meet the users specifications (i[4])
                if identcol[j] in x[4]:
                    okrows.append(j)
            statsx=[] #start a new list
            statsx.append(a[x[0]]+'-'+xvalstart[0]+'('+str(x[4])+')') #add the column header and filter
            for k in range(1,len(xvalstart)): #for all of the values
                if k in okrows: #if the value comes from a selected row, add it to the list
                    statsx.append(xvalstart[k])
                else: #otherwise, append a placeholder so that you can directly compare the indices to the list of y values
                    statsx.append('')
        else:
            statsx=[a[x[0]]+'-'+xvalstart[0]]
            statsx=statsx+xvalstart[1:] #if the user didn't choose to filter, just copy the list verbatim
        for y in yaxes: #same logic as above, but for y
            yvalstart=[]
            yvalstart=copycol(book,y[0],y[1]) #copy that column
            if y[3]==1: #if the user chose to filter based on a numerical value
                statsy=[] # create an output list
                statsy.append(a[y[0]]+'-'+yvalstart[0]+'('+y[4][0]+y[4][1]+')') #add the column header and filter
                statsy=statsy+conservsortfrominput(y[4],yvalstart[1:]) #sort based on the user's input (i[4]), conserving the index
            elif y[3]==2: #if the user chose to filter based on an experiment identifier
                headings=colheadingreadernum(book,y[0]) #look at the headers
                identcol=copycol(book,y[0],headings.index('Experiment Identifier')) #find the column for the experiment identifier
                okrows=[] #start an empty list
                for j in range(len(identcol)): #copy all of the row numbers that meet the users specifications (i[4])
                    if identcol[j] in y[4]:
                        okrows.append(j)
                statsy=[] #start a new list
                statsy.append(a[y[0]]+'-'+yvalstart[0]+'('+str(y[4])+')') #add the column header and filter
                for k in range(1,len(yvalstart)): #for all of the values
                    if k in okrows: #if the value comes from a selected row, add it to the list
                        statsy.append(yvalstart[k])
                    else:
                        statsy.append('') #conserve the index
            else:
                statsy=[a[y[0]]+'-'+yvalstart[0]]
                statsy=statsy+yvalstart[1:]  #if the user didn't choose to filter, just copy the list verbatim
            xfinal=[]
            xfinal.append(statsx[0]) #move the header (including filter if any)
            yfinal=[]
            yfinal.append([statsy[0],y[2]]) #pass the heading and whether or not the user wanted to filter AND graph
            for i in range(1,len(xvalstart)): #If BOTH the x and y parameters were not filtered out (ie still numbers), add to the final list
                if type(statsx[i])==float:
                    if type(statsy[i])==float:
                        xfinal.append(statsx[i])
                        yfinal.append(statsy[i])
                    else:
                        pass
                else:
                    pass
            statstorun.append([xfinal,yfinal]) #append the final list for each parameter to a master list
    return statstorun

def dothestuff(book,bookpath,writesheet,writesheetrows):
    c=arrangexy(book) #Pull the parameters to run statistics on
    writesheet.col(0).width=15000 #Make the first column wider to accomodate long parameter names
    row=writesheetrows+1 # set the starting row
    for i in c:
        slope,intercept,r_value,pvalue,std_err=stats.linregress(i[0][1:],i[1][1:]) #do a linear regression of the two lists
        xlabel=i[0][0] #otherwise, use the parameter name alone
        xtitlehalf=i[0][0]+' vs.'
        writesheet.write(row,0,'X='+i[0][0])
        ylabel=i[1][0][0]
        ytitlehalf=i[1][0][0]
        writesheet.write(row+1,0,'Y='+ytitlehalf)
        writesheet.write(row+2,0,'n='+str(len(i[0])-1)) #write the calculated parameters, including n...
        writesheet.write(row,1,'slope')
        writesheet.write(row+1,1, slope) #... slope of the best-fit-line
        writesheet.write(row,2,'intercept')
        writesheet.write(row+1,2,intercept)#... intercept of the best-fit-line
        writesheet.write(row,3,'r^2')
        writesheet.write(row+1,3,r_value**2) #...r^2 value
        writesheet.write(row,4,'p value')
        writesheet.write(row+1,4,pvalue) #... and p value
        if i[1][0][1]==1: #if the user chose to graph the results
            title=xtitlehalf+ytitlehalf #create a title
            writesheet.row(row).set_style(xlwt.easyxf('font:height 5000')) #make the row tall enough to accomodate it
            lista=np.array(i[0][1:]) #set the lists to a numpy array- makes the program less cranky (don't ask me why, it just does)
            listb=np.array(i[1][1:])
            graphxy(lista,listb,title,bookpath+str(c.index(i)),xlabel,ylabel,slope,intercept,r_value,pvalue,size=(480,480),marker='rx',markersize=8) #graph it---CUSTOMIZABLE HERE
            writesheet.insert_bitmap(bookpath+str(c.index(i))+'.bmp',row,6) #Put the figure into the excel sheet
            os.remove(bookpath+str(c.index(i))+'.png')
            os.remove(bookpath+str(c.index(i))+'.bmp')
        row=row+4 #move down 4 rows for the next parameter
    return row
