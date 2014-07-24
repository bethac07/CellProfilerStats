import xlrd
import xlwt
from xlutils.copy import copy
import matplotlib.pyplot as plt
from matplotlib import cm
from PIL import Image
import os
import easygui as eg
import csv
from scipy import stats
import textwrap
import numpy

def pythag(xb,xa,yb,ya): #the pythagorean theorem
    return (((xb-xa)**2)+((yb-ya)**2))**(0.5)

def unziprezip(inlist): 
    '''turns [[1,2,3],[a,b,c]] into [[1,a],[2,b],[3,c]]'''
    outlist=[]
    for i in range(len(inlist[0])):
        temp=[]
        for j in range(len(inlist)):
            temp.append(inlist[j][i])
        outlist.append(temp)
    return outlist
    
def deepunziprezip(inlist): 
    '''turns [[[1,2],[3,4]],[[5,6],[7,8]]] into [[[1,5],[2,6]],[[3,7],[4,8]]]''' 
    outlist=[]
    for i in range(len(inlist[0])):
        outtemp=[]
        length=[]
        for k in inlist[0]:
            length.append(len(k))
        for j in range(max(length)):
            intemp=[]
            for m in range(len(inlist)):
                try:
                    intemp.append(inlist[m][i][j])
                except:
                    pass
            outtemp.append(intemp)
        outlist.append(outtemp)
    return outlist

def partialcount(list,query):
    """totals instances of a query in all members of a list (ie 4 a's in a list of ['apple','banana'])"""
    count=0
    for i in list:
        subcount=i.count(query)
        count+=subcount
    return count

def copysheet(book,sheet,startrow=0,startcol=0,stoprow='a',stopcol='a'): 
    '''copy all the information in a sheet to a nested list (list of all rows(list of each column value inside the row))'''
    t=[]
    b=book.sheet_by_index(sheet)
    if stoprow=='a':
        stoprow=b.nrows
    if stopcol=='a':
        stopcol=b.ncols
    for i in range(startrow,stoprow):
        brow=[]
        for k in range(startcol,stopcol):
            brow.append(b.cell(i,k).value)
        t.append(brow)
    return t

def maxnumcols(inputlist): 
    '''calculate the maximum length of each sub-list of a larger list'''
    t=[0]
    for i in inputlist:
        if len(i) > t[0]:
            del t[0]
            t.append(len(i))
    return t[0]

def copycol(book,sheet,col,startrow=0): 
    '''copy a single column to a list-can just use sheet.col_values(x)'''
    b=book.sheet_by_index(sheet)
    #t=b.col_values(col)[startrow:]
    t=[]
    for i in range(startrow,b.nrows):
        t.append(b.cell(i,col).value)
    return t
    
def copycolname(book,sheet,col,startrow=0): 
    '''copy a single column to a list-can just use sheet.col_values(x)'''
    b=book.sheet_by_name(sheet)
    t=b.col_values(col)[startrow:]
    """t=[]
    for i in range(startrow,b.nrows):
        t.append(b.cell(i,col).value)"""
    return t

def writesheet(sheet,inputlist,startrow=0,startcol=0): 
    '''write a nested list to a worksheet, default starting from top left'''
    for i in range(len(inputlist)):
        for j in range(len(inputlist[i])):
            sheet.write(i+startrow,j+startcol,inputlist[i][j])
            
def writesheetspacing(sheet,inputlist,xspacing=1,yspacing=1,startrow=0,startcol=0): 
    '''write a list to a worksheet, default starting from top left'''
    for i in range(len(inputlist)):
        for j in range(len(inputlist[i])):
            sheet.write((i*yspacing)+startrow,(j*xspacing)+startcol,inputlist[i][j])

def writecol(sheet,inputlist,col,startrow=0): 
    '''write a column to a worksheet, default starting at the top'''
    for i in range(len(inputlist)):
        sheet.write(i+startrow,col,inputlist[i])

def stripheaderonce(t,delimiter='_'): 
    '''removes the header before and up to the first underscore (or whatever delimiter chosen) from a string'''
    if delimiter in t:
        t=t[t.index(delimiter)+1:]
    return t

def stripheaderagg(t,delimiter='_'): 
    '''removes the header before and up to the last underscore (or whatever delimiter chosen) from a string'''
    while delimiter in t:
        t=t[t.index(delimiter)+1:]
    return t
    
def middleofdelim(t,delimiter='_'):
    '''pulls any text between (the first two) instances of a delimiter'''
    if delimiter in t:
        t=t[t.index(delimiter)+1:]
    if delimiter in t:
        t=t[:t.index(delimiter)]
    return t

def stripheaderlist(t,delimiter='_'): 
    '''removes the header before and up to an underscore (or whatever delimiter chosen) from each string in a list'''
    s=[]
    for i in t:
        i=str(i)
        if delimiter in i:
            s.append(i[(i.index(delimiter)+1):])
        else:
            s.append(i)
    return s

def colheadingreadername(book,sheet): 
    '''read the first row of a sheet, ie the column headings-just sheet.row)values(0)'''
    b=book.sheet_by_name(sheet)
    """t=[]
    for i in range(b.ncols):
        t.append(b.cell(0,i).value)"""
    t=b.row_values(0)
    return t

def colheadingreadernum(book,sheet): 
    '''read the first row of a sheet, ie the column headings-just sheet.row)values(0)'''
    b=book.sheet_by_index(sheet)
    """t=[]
    for i in range(b.ncols):
        t.append(b.cell(0,i).value)"""
    t=b.row_values(0)
    return t

def readsheets1file(excelfile): 
    """return the stripped names of all the sheets in a workbook"""
    msheets=[]
    for sheet_name in excelfile.sheet_names():
        msheets.append(sheet_name)
    msheets=stripheaderlist(msheets)
    return msheets

def readsheets2files(excelfile1,excelfile2): 
    '''reads two workbooks and returns the stripped names of their component sheets'''
    msheets=[]
    asheets=[]
    for sheet_name in excelfile1.sheet_names():
        msheets.append(sheet_name)
    for sheet_name in excelfile2.sheet_names():
        asheets.append(sheet_name)
    msheets=stripheaderlist(msheets)
    asheets=stripheaderlist(asheets)
    return [msheets,asheets]


def sortfrominput(b,i): 
    '''take user input for a filter (ie x<=2) and turns it from a string into an equation,
    then passes ONLY the values from the list that satisfy that filter'''
    t=[]
    if b[0]=='<=':
        for j in i:
            if j<=(float(b[1])):
                t.append(j)
    elif b[0]=='<':
        for j in i:
            if j<(float(b[1])):
                t.append(j)
    elif b[0]=='==':
        for j in i:
            if j==(float(b[1])):
                t.append(j)
    elif b[0]=='>=':
        for j in i:
            if j>=(float(b[1])):
                t.append(j)
    elif b[0]=='>':
        for j in i:
            if j>(float(b[1])):
                t.append(j)
    elif b[0]=='!=':
        for j in i:
            if j!=(float(b[1])):
                t.append(j)
    elif b[0]=='Percentile<=':
        cutoff=stats.scoreatpercentile(i,float(b[1]))
        for j in i:
            if j<=cutoff:
                t.append(j)
    elif b[0]=='Percentile>=':
        cutoff=stats.scoreatpercentile(i,float(b[1]))
        for j in i:
            if j>=cutoff:
                t.append(j)
    else:
        pass
    return t

def sortandgiveindex(b,i): 
    '''take user input for a filter (ie x<=2) and turns it from a string into an equation,
    then passes the indices from the list that satisfy that filter'''
    t=[]
    if b[0]=='<=':
        for j in range(len(i)):
            if type(i[j])==float:
                if i[j]<=(float(b[1])):
                    t.append(j)
    elif b[0]=='<':
        for j in range(len(i)):
            if type(i[j])==float:
                if i[j]<(float(b[1])):
                    t.append(j)
    elif b[0]=='==':
        for j in range(len(i)):
            if type(i[j])==float:
                if i[j]==(float(b[1])):
                    t.append(j)
    elif b[0]=='>=':
        for j in range(len(i)):
            if type(i[j])==float:
                if i[j]>=(float(b[1])):
                    t.append(j)
    elif b[0]=='>':
        for j in range(len(i)):
            if type(i[j])==float:
                if i[j]>(float(b[1])):
                    t.append(j)
    elif b[0]=='!=':
        for j in range(len(i)):
            if type(i[j])==float:
                if i[j]!=(float(b[1])):
                    t.append(j)
    elif b[0]=='Percentile<=':
        cutoff=stats.scoreatpercentile(i[1:],float(b[1]))
        for j in range(len(i)):
            if type(i[j])==float:
                    if i[j]<=cutoff:
                        t.append(j)
    elif b[0]=='Percentile>=':
        for j in range(len(i)):
            cutoff=stats.scoreatpercentile(i[1:],float(b[1]))
            if type(i[j])==float:
                    if i[j]>=cutoff:
                        t.append(j)
    else:
        pass
    return t

def conservsortfrominput(b,i): 
    '''take user input for a filter (ie x<=2) and turns it from a string into an equation,
    then passes the a list of the same size where the value is retained if it passes 
    the filter, or is turned into '' if it does not'''
    t=[]
    if b[0]=='<=':
        for j in i:
            if j=='':
                t.append(j)
            elif j<=(float(b[1])):
                t.append(j)
            else:
                t.append('')
    elif b[0]=='<':
        for j in i:
            if j=='':
                t.append(j)
            elif j<(float(b[1])):
                t.append(j)
            else:
                t.append('')
    elif b[0]=='==':
        for j in i:
            if j=='':
                t.append(j)
            elif j==(float(b[1])):
                t.append(j)
            else:
                t.append('')
    elif b[0]=='>=':
        for j in i:
            if j=='':
                t.append(j)
            elif j>=(float(b[1])):
                t.append(j)
            else:
                t.append('')
    elif b[0]=='>':
        for j in i:
            if j=='':
                t.append(j)
            elif j>(float(b[1])):
                t.append(j)
            else:
                t.append('')
    elif b[0]=='!=':
        for j in i:
            if j=='':
                t.append(j)
            elif j!=(float(b[1])):
                t.append(j)
            else:
                t.append('')
    else:
        pass
    return t

def axisvaluefinder(columnheading):
    '''Function to try to clean up axes a bit for graphs'''
    if type(columnheading)!=str:
        columnheading=str(columnheading)
    if '_Area' in columnheading:
        value='Area (pixels^2)'
    elif 'Eccentricity' in columnheading:
        value='Eccentricity'
    elif 'Intensity' in columnheading:
        value=stripheaderlist([columnheading])[0]+'(arbitrary units)'
    elif 'Distance' in columnheading:
        value='Distance (pixels)'
    elif 'To' in columnheading:
        value='Distance (pixels)'
    elif 'Length' in columnheading:
        value='Length (pixels)'
    else:
        value=columnheading
    return value

def graphxy(listx,listy,title,saveas,xlabel,ylabel,slope,intercept,r_value,pvalue,size=(480,480),marker='rx',markersize=8):
    """Graph two parameters against each other and return a graph the linear regression, r value, and p value"""
    plt.ioff() #turn off interactive mode
    plt.figure() 
    plt.plot(listx, listy, marker,ms=markersize) #plot the figure
    plt.plot(listx,(slope*listx+intercept), '-k', linewidth=1) #plot the best-fit line calculated above
    plt.figtext(.15,.85,'y='+str(slope)+'*x+'+str(intercept)) #add the equation of the line
    plt.figtext(.15,.8,'r^2='+str(r_value**2)) #add the r2
    plt.figtext(.15,.75,'p value='+str(pvalue)) #add the p value
    plt.title(title) #add the title and axis labels
    plt.xlabel(axisvaluefinder(xlabel))
    plt.ylabel(axisvaluefinder(ylabel))
    plt.savefig(saveas+'.png') #save as a .png to then change to .bmp for the excel reader
    img=Image.open(saveas+'.png')
    size=size #shrink the size
    if len(img.split()) == 4: #if it's RGBA instead of RGB, get rid of the 4th channel
        # prevent IOError: cannot write mode RGBA as BMP
        r, g, b, a = img.split()
        img = Image.merge("RGB", (r, g, b))
        img.thumbnail(size) #shrink it and save it
        img.save(saveas+'.bmp')
    else:
        img.thumbnail(size) #if it's RGB, just shrink it and save it
        img.save(saveas+'.bmp')
        
def graphbubble(listofxs,listofys,listofsizes,listoflabels,saveas,xlabel,ylabel,sizelabel, log=False, mantitle=None,size=(480,480),markerlist=['o','^','s','p','h','d'],colorlist=['b','r','g','y','k','m','c'],savefiles=False,PDF=False):
    '''Function for making bubble charts in matplotlib (ie a scatter plot where size of the marker represents the
    value of a third parameter)'''
    minval=min(listofsizes[0])
    maxval=max(listofsizes[0])
    #fix THIS
    if len(listofsizes)>1:
        for subsizelist in listofsizes:
            if min(subsizelist)<minval:
                minval=min(subsizelist)
            if max(subsizelist)>maxval:
                maxval=max(subsizelist)
         #print minval,maxval
    adjsize=[]
    for subsizelist in listofsizes:
        subadjsize=[]
        for subsize in subsizelist:
            subadjsize.append(5+int(((subsize-minval)*75)/(maxval-minval)))
        adjsize.append(subadjsize)
        #print min(subadjsize),max(subadjsize)
    plt.ioff() #turn off interactive mode
    plt.figure() 
    ax=plt.subplot(111)
    for i in range(len(listofxs)):
        ax.scatter(listofxs[i],listofys[i],label=listoflabels[i],marker=markerlist[i%6],c=colorlist[i%7],alpha=0.3,s=adjsize[i])
    if mantitle==None:
        dashindex=xlabel.index('-')
        mantitle=xlabel+' vs. '+ylabel[dashindex+1:]+' vs. '+sizelabel[dashindex+1:] #add the title and axis labels
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    # Put a legend to the right of the current axis
    leg=ax.legend(numpoints=1,loc='center left', bbox_to_anchor=(1, 0.5))
    #for t in leg.get_texts():
    #    t.set_fontsize('xx-small')
    if log!=False:
        if 'y' in log:
            plt.yscale('log')
        if 'x' in log:
            plt.xscale('log')
    plt.xlabel(axisvaluefinder(xlabel))
    plt.ylabel(axisvaluefinder(ylabel))
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        if mantitle==None:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+xlabel+' vs. '+ylabel[dashindex+1:]+' vs. '+sizelabel[dashindex+1:]+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],xlabel+' vs. '+ylabel[dashindex+1:]+' vs. '+sizelabel[dashindex+1:]+'.svg')
        else:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+mantitle+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename,transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)  

def graphscatter(listofxs,listofys,listoflabels,saveas,xlabel,ylabel,mantitle=None,size=(480,480),markersize=6,markerlist=['o','^','s','p','H','D'],colorlist=['b','r','g','y','k','m','c'],savefiles=False,PDF=False,log=False):
    '''function for graphing scatter plots in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    ax=plt.subplot(111)
    ax.set_rasterization_zorder(1)
    for i in range(len(listofxs)):
        #print i%6,i%7
        ax.plot(listofxs[i],listofys[i],label=listoflabels[i],marker=markerlist[i%6],color=colorlist[i%7],alpha=0.3,ms=markersize,linestyle='None',zorder=0)
    if mantitle==None:
        dashindex=xlabel.index('-')
        mantitle=xlabel+' vs. '+ylabel[dashindex+1:] #add the title and axis labels
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    # Put a legend to the right of the current axis
    leg=ax.legend(numpoints=1,loc='center left', bbox_to_anchor=(1, 0.5))
    for t in leg.get_texts():
        t.set_fontsize('xx-small')
    plt.xlabel(axisvaluefinder(xlabel))
    plt.ylabel(axisvaluefinder(ylabel))
    #plt.margins(tight=False)
    if log!=False:
        if 'y' in log:
            plt.yscale('log')
        if 'x' in log:
            plt.xscale('log')
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        if mantitle==None:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+xlabel+' vs. '+ylabel[dashindex+1:]+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],xlabel+' vs. '+ylabel[dashindex+1:]+'.svg')
        else:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+mantitle+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)  

def graphswarm(listofvalues,listoflabels,saveas,ylabel,mantitle=None,size=(480,480),markersize=3,savefiles=False,PDF=False,log=False,alpha=0.3):
    '''function for graphing scatter plots in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    ax=plt.subplot(111)
    ax.set_rasterization_zorder(1)
    for i in range(len(listofvalues)):
        if len(listofvalues[i])>0:
            listofxs=list(numpy.random.permutation(numpy.arange(i+0.7,i+1.3,(0.6/len(listofvalues[i])))))
            if len(listofvalues[i])<len(listofxs):
                listofxs=listofxs[:-(len(listofxs)-len(listofvalues[i]))]
            elif len(listofvalues[i])>len(listofxs):
                for extras in range(len(listofvalues[i])-len(listofxs)):
                    listofxs.append(i+1)
            #print len(listofxs),len(listofvalues[i])
            ax.plot(listofxs,listofvalues[i],label=listoflabels[i],marker='o',color='k',alpha=alpha,ms=markersize,linestyle='None',zorder=0)
            ax.plot([i+0.6,i+1.4],[numpy.median(listofvalues[i]),numpy.median(listofvalues[i])],color='r')
    if log==True:
        plt.yscale('log')
    if mantitle==None:
        mantitle=ylabel #add the title and axis labels
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.95, box.height])
    # Put a legend to the right of the current axis
    plt.ylabel(axisvaluefinder(ylabel))
    oldymin,oldymax=plt.ylim()
    plt.ylim(0.9*oldymin,1.1*oldymax)
    ax.set_xticks(range(1,len(listoflabels)+1))
    ax.set_xticklabels(listoflabels,fontsize='xx-small',rotation=45)
    #plt.margins(tight=False)
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        if mantitle==None:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+ylabel+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],ylabel+'.svg')
        else:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+mantitle+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)  

def graphsubgroupswarm(listofinvalues,listofoutvalues,listoflabels,saveas,ylabel,mantitle=None,size=(480,480),markersize=3,savefiles=False,PDF=False,log=False):
    '''function for graphing scatter plots in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    ax=plt.subplot(111)
    ax.set_rasterization_zorder(1)
    labelcopy=[]
    #ttest=stats.ttest_ind(chromdict[5],chromdict['all'],equal_var=False)
    for i in range(len(listofinvalues)):
        labelcopy.append(listoflabels[i])
        if len(listofoutvalues[i])>0:
            listofoutxs=list(numpy.random.permutation(numpy.arange(i+0.7,i+1.3,(0.6/len(listofoutvalues[i])))))
            if len(listofoutvalues[i])<len(listofoutxs):
                listofoutxs=listofoutxs[:-(len(listofoutxs)-len(listofoutvalues[i]))]
            elif len(listofoutvalues[i])>len(listofoutxs):
                for extras in range(len(listofoutvalues[i])-len(listofoutxs)):
                    listofoutxs.append(i+1)
            #print len(listofxs),len(listofvalues[i])
            ax.plot(listofoutxs,listofoutvalues[i],label=listoflabels[i],marker='o',color='k',alpha=0.3,ms=markersize,linestyle='None',zorder=0)
            ax.plot([i+0.6,i+1.4],[numpy.median(listofoutvalues[i]),numpy.median(listofoutvalues[i])],color='k',linewidth=2)
        else:
            labelcopy[i]+=('\n n.d.')
        if len(listofinvalues[i])>0:
            listofinxs=list(numpy.random.permutation(numpy.arange(i+0.7,i+1.3,(0.6/len(listofinvalues[i])))))
            if len(listofinvalues[i])<len(listofinxs):
                listofinxs=listofinxs[:-(len(listofinxs)-len(listofinvalues[i]))]
            elif len(listofinvalues[i])>len(listofinxs):
                for extras in range(len(listofinvalues[i])-len(listofinxs)):
                    listofinxs.append(i+1)
            #print len(listofxs),len(listofvalues[i])
            ax.plot(listofinxs,listofinvalues[i],label=listoflabels[i],marker='o',color='r',alpha=0.3,ms=markersize,linestyle='None',zorder=0)
            ax.plot([i+0.6,i+1.4],[numpy.median(listofinvalues[i]),numpy.median(listofinvalues[i])],color='r',linewidth=2)
            if len(listofoutvalues[i])>0:
                 pvalue=stats.ttest_ind(listofinvalues[i],listofoutvalues[i],equal_var=False)[1]
                 if pvalue<0.001:
                     labelcopy[i]+='\n p<0.001'
                 elif pvalue<0.01:
                     labelcopy[i]+='\n p<0.01'
                 elif pvalue<0.05:
                     labelcopy[i]+='\n p<0.05'
                 else:
                     labelcopy[i]+='\n n.s.'
        else:
           labelcopy[i]+=('\n n.d.') 
    if log==True:
        plt.yscale('log')
    if mantitle==None:
        mantitle=ylabel #add the title and axis labels
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.95, box.height])
    # Put a legend to the right of the current axis
    plt.ylabel(axisvaluefinder(ylabel))
    oldymin,oldymax=plt.ylim()
    plt.ylim(0.9*oldymin,1.1*oldymax)
    ax.set_xticks(range(1,len(labelcopy)+1))
    ax.set_xticklabels(labelcopy,fontsize='xx-small',rotation=45)
    #plt.margins(tight=False)
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        if mantitle==None:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+ylabel+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],ylabel+'.svg')
        else:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+mantitle+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)

def graphscumhist(listofxs,listoflabels,saveas,xlabel,mantitle=None,size=(480,480),markersize=6,markerlist=['o','^','s','p','H','D'],colorlist=['b','r','g','y','k','m','c','Sienna','DeepPink','Lime'],savefiles=False,PDF=False,log=False):
    '''function for graphing scatter plots in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    listoflabelscopy=[]
    for j in range(len(listoflabels)):
        listoflabelscopy.append("\n".join(textwrap.wrap(listoflabels[j]+' n='+str(len(listofxs[j])))))
    ax=plt.subplot(111)
    ax.set_rasterization_zorder(1)
    for i in range(len(listofxs)):
        #print i%6,i%7
        try:
            percents=range(0,100)
            percentvals=[]
            for j in percents:
                percentvals.append(stats.scoreatpercentile(listofxs[i],j))
            ax.plot(percentvals,percents,label=listoflabelscopy[i],color=colorlist[i%10],zorder=0)
        except:
            pass
    if log==True:
        plt.xscale('log')
    if mantitle==None:
        mantitle='Cumulative Frequency of '+xlabel #add the title and axis labels
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    # Put a legend to the right of the current axis
    leg=ax.legend(numpoints=1,loc='center left', bbox_to_anchor=(1, 0.5))
    for t in leg.get_texts():
        t.set_fontsize('xx-small')
    plt.xlabel(xlabel)
    plt.ylabel('Cumulative Frequency')
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        if mantitle==None:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+xlabel+' CumulativeFreq.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],xlabel+' CumulativeFreq.svg')
        else:
            if len(listoflabels)==1:
                filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+mantitle+'.svg')
            else:
                filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)  

def graphmsds(listoftimepoints,listofmsds,saveas,labels=False,xunits='sec',yunits='(um)',mantitle='Per-Track MSDs',size=(300,300),savefiles=False,PDF=False,each=False,colorlist=['b','r','g','y','k','m','c','Sienna','DeepPink','Lime']):
    '''function for graphing MSD plots in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    ax=plt.subplot(111)
    ax.set_rasterization_zorder(1)
    if each==False:
        for i in range(len(listoftimepoints)):
            if labels==False:
                ax.plot(listoftimepoints[i],listofmsds[i],zorder=0)
            else:
                ax.plot(listoftimepoints[i],listofmsds[i],zorder=0,label=labels[i],color=colorlist[i])
    else:
        for i in range(len(listoftimepoints)):
            unzippeddata=unziprezip(listoftimepoints[i])
            ax.plot(unzippeddata[0],unzippeddata[1],zorder=0)
    if labels!=False:
        box = ax.get_position()
        ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
        # Put a legend to the right of the current axis
        leg=ax.legend(numpoints=1,loc='center left', bbox_to_anchor=(1, 0.5))
        for t in leg.get_texts():
            t.set_fontsize('xx-small')
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    plt.xlabel('Time ('+xunits+')')
    plt.ylabel('MSD '+yunits[:-1]+'^2)')
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)
    plt.close()

def graphtracksinacell(trackdict,saveas,mantitle,size=(480,480),colorlist=['LightSkyBlue','DeepSkyBlue','Blue','Black'],savefiles=False,PDF=False):
    '''function for graphing cumulative track movements in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    ax=plt.subplot(111)
    ax.set_rasterization_zorder(1)
    for i in trackdict:
        ax.plot(i[1],i[2],color=colorlist[i[0]],zorder=0)
    if mantitle==None:
        mantitle='Integrated Distance of Tracks' #add the title and axis labels
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    plt.xlabel('Frame #')
    plt.ylabel('Integrated distance traveled (pixels)')
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()

def graphhist(valuelist,title,saveas,xlabel,ylabel='This axis is wrong',size=(480,480)):
    plt.ioff() #turn off interactive mode to run more quickly
    plt.figure() #start a new figure
    plt.hist((valuelist[1:]),20,normed=True) #create a 20-binned histogram <---------More parameters can be changed here if desired
    plt.title(title) #title the histogram with what it is
    plt.xlabel(axisvaluefinder(xlabel))
    plt.ylabel(axisvaluefinder(ylabel))
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    size=480,480 #shrink the size
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    img.thumbnail(size) #if it's RGB, just shrink it and save it
    img.save(file_out)

def graph2hists(lista,listb, labela, labelb,title,saveas,statvalue,pvalue,xlabel,ylabel='This axis is wrong',size=(480,480)):
    plt.ioff() #turn off interactive mode to run more quickly
    plt.figure() #start a new figure
    n,bins,patches=plt.hist([lista,listb],label=[labela,labelb],normed=True,alpha=0.8,bins=20) #create a 20-binned histogram <---------More parameters can be changed here if desired
    plt.title(title) #title the histogram with what it is
    plt.xlabel(axisvaluefinder(xlabel))
    plt.ylabel(axisvaluefinder(ylabel))
    plt.figtext(.15,.85,'u value='+str(statvalue))
    plt.figtext(.15,.80,'p value='+str(pvalue))
    plt.legend()
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    size=480,480 #shrink the size
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    img.thumbnail(size) #if it's RGB, just shrink it and save it
    img.save(file_out)
    
def graphmanyhists(listoflists, listoflabels,title,saveas,xlabel='',ylabel='% of values',normed=True, size=(300,300),colorlist=['b','r','g','y','k','m','c','Sienna','DeepPink','Lime'],PDF=False,log=False):
    plt.ioff() #turn off interactive mode to run more quickly
    plt.figure() #start a new figure
    listoflabelscopy=[]
    if len(listoflists)>16:
        for k in range(16,len(listoflists)):
            colorlist.append(colorlist[k%16])
    else:
        colorlist=colorlist[0:len(listoflists)]
    for j in range(len(listoflabels)):
        listoflabelscopy.append(listoflabels[j]+' n='+str(len(listoflists[j])))
    if normed==True:
        weightlist=[]
        for eachlist in listoflists:
            setlen=len(eachlist)
            weightlist.append(setlen*[100.0/float(setlen)])
    ax=plt.subplot(111)
    if normed==False:
        plt.hist(listoflists,label=listoflabelscopy,cumulative=False, alpha=0.8,bins=20,log=log,color=colorlist) #create a 20-binned histogram <---------More parameters can be changed here if desired
    else:
        plt.hist(listoflists,label=listoflabelscopy,cumulative=False, weights=weightlist,alpha=0.8,bins=20,log=log,color=colorlist) #create a 20-binned histogram <---------More parameters can be changed here if desired
    # Shink current axis by 20%
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    # Put a legend to the right of the current axis
    formtitle="\n".join(textwrap.wrap(title))
    plt.title(formtitle,stretch='ultra-condensed')#title the histogram with what it is
    plt.xlabel(axisvaluefinder(xlabel))
    plt.ylabel(axisvaluefinder(ylabel))
    leg=plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    for t in leg.get_texts():
        t.set_fontsize('xx-small')
    #plt.show()
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    #print saveas
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)   

def graphenrich(listofnums,listofdenoms,listoflabels,saveas,mantitle,size=(480,480),markersize=6,markerlist=['o','^','s','p','H','D'],colorlist=['b','r','g','y','k','m','c','Sienna','DeepPink','Lime'],savefiles=False,PDF=False):
    '''function for graphing stuff in matplotlib'''   
    plt.ioff() #turn off interactive mode
    plt.figure() 
    listoflabelscopy=[]
    for j in range(len(listoflabels)):
        listoflabelscopy.append(listoflabels[j]+' n='+str(len(listofnums[j]))+'/'+str(len(listofdenoms[j])))
    ax=plt.subplot(111)
    for i in range(len(listofnums)):
        #print i%6,i%7
        percents=range(0,101,5)
        percentvals=[]
        for j in percents:
            percentvals.append(stats.scoreatpercentile(listofnums[i],j)/stats.scoreatpercentile(listofdenoms[i],j))    
        ax.plot(percents,percentvals,label=listoflabelscopy[i],color=colorlist[i%10])
    formtitle="\n".join(textwrap.wrap(mantitle))
    plt.title(formtitle,stretch='ultra-condensed')
    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    # Put a legend to the right of the current axis
    leg=ax.legend(numpoints=1,loc='center left', bbox_to_anchor=(1, 0.5))
    for t in leg.get_texts():
        t.set_fontsize('xx-small')
    plt.xlabel('Percentile')
    plt.ylabel('Relative Value')
    if PDF!=False:
        plt.savefig(PDF, format='pdf')
    if savefiles==True:
        if len(listoflabels)==1:
            filename=os.path.join(os.path.split(saveas)[0],listoflabels[0]+'-'+mantitle+'.svg')
        else:
            filename=os.path.join(os.path.split(saveas)[0],mantitle+'.svg')
        plt.savefig(filename, transparent=True)
    plt.savefig(saveas+'.png') #save and close the .png
    plt.close()
    img = Image.open(saveas+'.png') #open the .png in PIL to convert it to a .bmp
    file_out = saveas+'.bmp' #the Excel reader only will use .bmp
    if img.mode!= 'RGB':
        img=img.convert('RGB')
    res=img.resize(size) #if it's RGB, just shrink it and save it
    res.save(file_out)  


def runSOMEFileIO():
    '''template for file I/O'''
    aa=eg.fileopenbox(msg='Choose input Excel file',default='*.xls') #select a file, open it
    book=xlrd.open_workbook(aa)
    bookpath=os.path.dirname(aa)
    sheetnames=readsheets1file(book) #read the sheet names
    writesheetrows=0 #set initial value of which row to write in as 0- is modified below if the user chooses to add to an existing sheet
    reuse=eg.indexbox('Where to save the output?',choices=['Add to a new sheet in the input file',
                                                           'Add to an existing sheet in the input file','Create a new file', 
                                                           'Add to a new sheet in another file','Add to an existing sheet in another file'])
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
    #dothestuff(book,bookpath,writesheet,writesheetrows)
    if reuse==0 or reuse==1: #if we're saving the input file, save it with the input name
        writebook.save(aa)
    elif reuse==2: #if we're saving a new file, save it with the name the user input
        writebook.save(w)
    elif reuse==3 or reuse==4: #if we're saving an old file, save it with it's original name
        writebook.save(o)

def arrangedivsin(book,a0):
    xaxes=a0 
    statstorun=[]
    sh=readsheets1file(book)
    for x in xaxes: #for x (or filtered and unfiltered if the user so chose)
        #print x
        xvalstart=[]
        xvalstart=copycol(book,x[0],x[1]) #copy that column
        if x[2]==1: #if the user chose to filter based on a numerical value
            statsx=[] # create an output list
            statsx.append(sh[x[0]]+'-'+xvalstart[0]+'('+x[3][0]+x[3][1]+')') #add the column header and filter
            statsx=statsx+sortfrominput(x[3],xvalstart[1:]) #sort based on the user's input (i[4])
            statstorun.append(statsx)
        elif x[2]==2:
            for relativesort in x[3]:
                #print relativesort
                if len(relativesort)==3: #if we're filtering based on another measure of that parameter (including # of children)
                    childcol=copycol(book,relativesort[0],relativesort[1])
                    useindices=sortandgiveindex(relativesort[2],childcol)
                    statsx=[sh[x[0]]+'-'+xvalstart[0]+'(have'+relativesort[2][0]+relativesort[2][1]+' '+str(childcol[0])+')']
                    for indtouse in useindices:
                        statsx.append(xvalstart[indtouse])
                    statstorun.append(statsx)
                elif len(relativesort)==5: #if we're filtering based on a parental factor
                    parentsheet=book.sheet_by_index(relativesort[2])
                    statsx=[sh[x[0]]+'-'+xvalstart[0]+'(have parent '+sh[relativesort[2]]+' whose '+str(parentsheet.cell(0,relativesort[3]).value)+' is ' +relativesort[4][0]+relativesort[4][1]+')']
                    statimagecol=copycol(book,relativesort[0],0)
                    findparentcol=copycol(book,relativesort[0],relativesort[1])
                    parentvaldict={}
                    for parentrow in range(1,parentsheet.nrows):
                        parrowvals=parentsheet.row_values(parentrow)
                        parentvaldict[(parrowvals[0],parrowvals[1])]=[parrowvals[relativesort[3]]]
                    for childrow in range(1,len(statimagecol)):
                        parentval=parentvaldict[(statimagecol[childrow],findparentcol[childrow])]
                        if len(sortfrominput(relativesort[4],parentval))!=0:
                            statsx.append(xvalstart[childrow])
                    statstorun.append(statsx)
                elif len(relativesort)==6: #if we're filtering based on having children that match a certain measurement
                    childsheet=book.sheet_by_index(relativesort[2])
                    statsx=[sh[x[0]]+'-'+xvalstart[0]+'(have child '+sh[relativesort[2]]+' whose '+str(childsheet.cell(0,relativesort[4]).value)+' is ' +relativesort[5][0]+relativesort[5][1]+')']
                    paramimagecol=copycol(book,relativesort[0],0)
                    paramidentcol=copycol(book,relativesort[0],1)
                    childvaldict={}
                    for childrow in range(1,childsheet.nrows):
                        childrowvals=childsheet.row_values(childrow)
                        if (childrowvals[0],childrowvals[relativesort[3]]) not in childvaldict.keys():
                            childvaldict[(childrowvals[0],childrowvals[relativesort[3]])]=[childrowvals[relativesort[4]]]
                        else:
                            childvaldict[(childrowvals[0],childrowvals[relativesort[3]])].append(childrowvals[relativesort[4]])
                    for paramrow in range(1,len(paramimagecol)):
                        checkifindict=(paramimagecol[paramrow],paramidentcol[paramrow])
                        if checkifindict in childvaldict.keys():
                            childval=childvaldict[checkifindict]
                            if len(sortfrominput(relativesort[5],childval))!=0:
                                statsx.append(xvalstart[paramrow])
                    statstorun.append(statsx)
                
        else:
            statsx=[sh[x[0]]+'-'+xvalstart[0]]
            statsx=statsx+xvalstart[1:] #if the user didn't choose to filter, just copy the list verbatim
            statstorun.append(statsx)
    return statstorun

def findexcel(directory): 
    '''Finds all files with a .csv extension in a given directory'''
    if not os.path.exists(directory): #Warn if directory doesn't exist
        print 'error: no such directory'
    else: #Create a list of the files
        t=[]
        for i in os.listdir(directory): #for all the files in the directory...
            if '.xls' in i: #if they have a .xls extension...
                j=os.path.join(directory,i) #create a full directory+filename string
                t.append([i,j]) #move to master list
            else: pass        
    return t

def csvtoexcel(csvin): 
    """Turns a .csv file into a single .xls file
    Returns the book and sheet name of that file to any application that calls it"""
    w=xlwt.Workbook() #Create new excel file
    folder,filelist=os.path.split(csvin)
    c=len(filelist) #see how long the filename is- for removing the .csv extension below
    while c>=35: #Sheet names can only be 31 characters long- prompt user if shorter sheet name required
        filelist=(eg.enterbox('Filename '+filelist+' is too long- enter a shorter one')+'.csv')
        c=len(filelist)
    b=w.add_sheet(filelist[:-4]) #name each sheet according to the name of it's .csv file
    a=csv.reader(open(csvin)) #Read the file
    linecount=0
    for i in a:
        linecount+=1
    if linecount>65535:
        broken=True
    else:
        broken=False
    count=0
    for i in a: #for all rows of information
        for j in range(len(i)): #for each piece of data in the row
            try:
                b.write(count,j,float(i[j]))
            except:
                b.write(count,j,i[j])
        count+=1
    k=os.path.join(folder,filelist[:-3]+'xls') #Save the excel file
    w.save(k)
    return k,filelist[:-4],broken
