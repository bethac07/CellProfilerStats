"""CSV Combiner- part of CellProfiler Stats- by Beth Cimini Sept 2010

    Takes the directory of .csv files created by the measurement algorithms of CellProfiler and combines them into
    a single Excel file, adding a nearest neighbor analysis for up to four types of objects if requested.
    
    TO FIX/Add Dec-21-11:
        Fix nearest neighbor to only look in same nucleus DONE 1/3/12
        Add %ABs of A's (per nucleus-children) DONE 1/7/12
        Add %Area/(Nuclear/otherobject) area and %II/(Nuclear/otherobject) II DONE 1/7/12
    Will need to change all the defaults for batch mode DONE 1/7/12"""

#from copy import deepcopy
import xlrd
import xlwt
from xlutils.copy import copy
import os
import csv
import math
import easygui as eg
import shelve
import numpy
import HandyXLModules as HXM


def findcsv(directory): #Finds all files with a .csv extension in a given directory
    if not os.path.exists(directory): #Warn if directory doesn't exist
        print 'error: no such directory'
    else: #Create a list of the files
        t=[]
        for i in os.listdir(directory): #for all the files in the directory...
            if '.csv' in i: #if they have a .csv extension...
                j=os.path.join(directory,i) #create a full directory+filename string
                t.append([i,j]) #move to master list
            else: pass        
    return t
 

def combinecsv(directory,outname): #Combines all .csv files into a single .xls file
    w=xlwt.Workbook() #Create new excel file
    filelist=findcsv(directory)
    shortfilelist=[]
    for y in filelist: #for each csv file...
        shortfilelist.append(y[0][:-4])
        c=len(y[0]) #see how long the filename is- for removing the .csv extension below
        if c>=35: #Sheet names can only be 31 characters long- prompt user if shorter sheet name required
            y[0]=(eg.enterbox('Filename '+y[0]+' is too long- enter a shorter one')+'.csv')
            c=len(y[0])
        b=w.add_sheet(y[0][0:c-4],cell_overwrite_ok=True) #name each sheet according to the name of it's .csv file
        a=csv.reader(open(y[1])) #Read the file
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
    k=os.path.join(directory,outname+'temp.xls') #Save the excel file
    w.save(k)
    return k,shortfilelist

def pullallvals(wb,sheetname,values,parentvalue):
    #pulls a value or values, returns a dictionary where the keys are a tuple of image and nucleus numbers and 
    #the value is a list of lists of the row and the value(s) for each object
    sheet=wb.sheet_by_name(sheetname)
    nucdic={}
    headings=HXM.colheadingreadername(wb,sheetname)
    imageindex=headings.index('ImageNumber')
    nucindex=headings.index(parentvalue)
    #nucindex=headings.index('Parent_Nuclei')
    valindex=[]
    for i in values:
        valindex.append(headings.index(i))
    for i in range(1,sheet.nrows):
        b=[i]
        for j in valindex:
            b.append(sheet.cell(i,j).value)
        if (sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value) not in nucdic.keys():
            nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]=[b]
        else:
            nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]+=[b]
    return nucdic

def makenormfactordiv(wb,sheetname,value):
    sheet=wb.sheet_by_name(sheetname)
    nucdic={}
    headings=HXM.colheadingreadername(wb,sheetname)
    imageindex=headings.index('ImageNumber')
    nucindex=headings.index('ObjectNumber')
    valindex=headings.index(value)
    masterlist=[]
    for i in range(1,sheet.nrows):
        b=sheet.cell(i,valindex).value
        masterlist.append(b)
        if (sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value) not in nucdic.keys():
            nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]=[b]
        else:
            nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]+=[b]
    normdic={}
    normdic['name']=HXM.stripheaderonce(sheetname)+'-'+HXM.stripheaderonce(value)
    mastermean=numpy.average(masterlist)
    for i in nucdic:
        normdic[i]=numpy.average(nucdic[i])/mastermean
    return normdic            
    
def makenormfactorsub(wb,sheetname,value):
    sheet=wb.sheet_by_name(sheetname)
    nucdic={}
    headings=HXM.colheadingreadername(wb,sheetname)
    imageindex=headings.index('ImageNumber')
    nucindex=headings.index('ObjectNumber')
    valindex=headings.index(value)
    for i in range(1,sheet.nrows):
        b=sheet.cell(i,valindex).value
        if (sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value) not in nucdic.keys():
            nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]=[b]
        else:
            nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]+=[b]
    normdic={}
    normdic['name']=HXM.stripheaderonce(sheetname)+'-'+HXM.stripheaderonce(value)
    for i in nucdic:
        normdic[i]=numpy.average(nucdic[i])
    return normdic 

def maketotal(wb,sheetname,value,parentvalue):
    sheet=wb.sheet_by_name(sheetname)
    nucdic={}
    headings=HXM.colheadingreadername(wb,sheetname)
    imageindex=headings.index('ImageNumber')
    nucindex=headings.index('Parent_'+parentvalue)
    valindex=headings.index(value)
    for i in range(1,sheet.nrows):
        b=sheet.cell(i,valindex).value
        try:
            b=float(b)
            if (sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value) not in nucdic.keys():
                nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]=b
            else:
                nucdic[(sheet.cell(i,imageindex).value,sheet.cell(i,nucindex).value)]+=b
        except:
            pass
    return nucdic    
    
def pythag2(xb,xa,yb,ya): #the pythagorean theorem
    return (((xb-xa)**2)+((yb-ya)**2))**(0.5)

def compare2dicts(dict1,dict2,dist=500):
    outdict={}
    for i in dict1.keys():
        if i in dict2.keys():
            dict2list=[]
            for j in dict2[i]:
                dict2list.append((j[1],j[2]))
            for k in dict1[i]:
                minval=dist*math.sqrt(2)
                for m in dict2list:
                    if abs(k[1]-m[0])<dist:
                        if abs(k[2]-m[1])<dist:
                            pythagval=pythag2(k[1],m[0],k[2],m[1])
                            if pythagval<minval:
                                minval=pythagval
                outdict[k[0]]=minval
        else:
            for j in dict1[i]:
                outdict[j[0]]=dist*math.sqrt(2)
    return outdict

def calcnearestneighbor(directory,csvoutname,default=None):
    #Can compare up to 4 objects (par1, par2, par3, par4)-
    #make sure the first object is par1, second is par2 etc if you're only comparing 2 or 3 objects

    #If you want to change the defaults for compare objects (either the objects or the names), the "def" above is the place to do it

    q,parlist=combinecsv(directory,csvoutname) #Combines all the CSVs- can be done separately if needed
    strippedparlist=[]
    for i in parlist:
        strippedparlist.append(i[i.index('_')+1:])
    rb=xlrd.open_workbook(q) #Open the excel file
    wb=copy(rb)
    sheetlist=HXM.readsheets1file(rb)
    headingdict={}
    for i in range(len(strippedparlist)):
        headingdict[strippedparlist[i]]=HXM.colheadingreadername(rb,parlist[i])
    numcoldict={}
    for i in range(len(sheetlist)):
        checksheetcol=rb.sheet_by_index(i)
        numcoldict[i]=checksheetcol.ncols
 
   
    if default==None:
        prevdef=shelve.open(os.path.join(os.curdir,'csvshelf'),writeback=True)
        usedef=eg.ynbox(msg='Do you want to use a default?')
        if usedef:
            default=eg.choicebox(msg='Pick the default you wish to use', choices=prevdef.keys())
            tocompinit,dist,kidanalysis,normanalysis,fractanalysis,waveratio=prevdef[default]
        
        else:
            if eg.ynbox(msg='Do you want to compare the locations of different objects?'):
                tocompinit=[]
                tocompinitchoices=eg.multchoicebox(msg='Pick the items whose locations you want to compare',choices=strippedparlist)
                allparentdic={}
                for eachchoice in tocompinitchoices:
                    for eachheading in headingdict[eachchoice]:
                        if 'Parent' in eachheading:
                            if eachheading not in allparentdic.keys():
                                allparentdic[eachheading]=[eachchoice]
                            else:
                                allparentdic[eachheading].append(eachchoice)
                for eachkey in allparentdic.keys():
                    if len(allparentdic[eachkey])==len(tocompinitchoices):
                        tocompinit=(tocompinitchoices,eachkey)
                if len(tocompinit)<=1:
                    tocompinit=False
                else:
                    dist=float(eg.enterbox(msg='Enter the maximum distance (in pixels) to look for nearby objects'))
            else:
                tocompinit=False
                dist=False
        
            if eg.ynbox(msg='Do you want to compare the number of 2 different children objects in the same parent (ie in a given nucleus, what is the ratio of telomeres to centromeres)?'):
                have2kids=[]
                allkids=[]
                readablekids=[]
                for k in headingdict.keys():
                    numkids=HXM.partialcount(headingdict[k],'Children')
                    if numkids>=2:
                        have2kids.append(k)
                        for j in headingdict[k]:
                            if 'Children' in j:
                                allkids.append((k,j))
                                readablekids.append(HXM.stripheaderonce(k)+' - '+HXM.middleofdelim(j))
                if len(have2kids)==0:
                    kidanalysis=False
                else:
                    numeratorkids=eg.multchoicebox(msg='Pick all the NUMERATOR children objects', choices=readablekids)
                    if len(numeratorkids)==0:
                        kidanalysis=False
                    else:
                        kidanalysis=[]
                        for num in numeratorkids:
                            subkidlist=[]
                            for subnum in readablekids:
                                if num[:num.index('-')]==subnum[:num.index('-')]:
                                    subkidlist.append(subnum)
                            denomkids=eg.multchoicebox(msg='Pick all the DENOMINATOR children objects for '+num,choices=subkidlist)
                            if len(denomkids)==0:
                                pass
                            else:
                                for subdenom in denomkids:
                                    kidanalysis.append((allkids[readablekids.index(num)],allkids[readablekids.index(subdenom)]))
                    
                    #print kidanalysis
                    if len(kidanalysis)==0:
                        kidanalysis=False
            else:
                kidanalysis=False
        
                
            if eg.ynbox(msg='Do you want to normalize any area or intensity measurements to the parental value?'):
                normanalysis=[]
                allaandi=[]
                readableaandi=[]
                for k in headingdict.keys():
                    for heads in headingdict[k]:
                        if '_Area' in heads:
                            allaandi.append((k,heads))
                            readableaandi.append(HXM.stripheaderonce(k)+' - '+HXM.stripheaderonce(heads))
                        elif 'Intensity' in heads:
                            allaandi.append((k,heads))
                            readableaandi.append(HXM.stripheaderonce(k)+' - '+HXM.stripheaderonce(heads))
                tobenormed=eg.multchoicebox(msg='Pick all the objects TO BE normalized', choices=readableaandi)
                if len(tobenormed)==0:
                    normanalysis=False
                else:
                    for sinnorm in tobenormed:
                        normers=eg.multchoicebox(msg='Pick all the objects to normalize '+sinnorm+' BY',choices=readableaandi)
                        if len(normers)==0:
                            pass
                        else:
                            for eachnorm in normers:
                                subordiv=eg.buttonbox(msg='Should division or subtraction be used to normalize '+sinnorm+' by '+eachnorm+'?',choices=('Division','Subtraction'))
                                normanalysis.append((allaandi[readableaandi.index(sinnorm)],allaandi[readableaandi.index(eachnorm)],subordiv))
                #print normanalysis
                #print (type(normanalysis))
                if len(normanalysis)==0:
                    normanalysis=False
            else:
                normanalysis=False
                
            if eg.ynbox(msg='Do you want to compare the fraction of any TOTAL child measurements to their parent measurements (ie, % of a nucleus covered by telomeres)?'):
                fractanalysis=[]
                have2kids=[]
                allkids=[]
                readablekids=[]
                for k in headingdict.keys():
                    for j in headingdict[k]:
                        if 'Children' in j:
                            allkids.append((k,j))
                            readablekids.append(HXM.stripheaderonce(k)+' - '+HXM.middleofdelim(j))
                numeratorkids=eg.multchoicebox(msg='Pick all the children objects', choices=readablekids)
                if len(numeratorkids)==0:
                    fractanalysis=False
                else:
                    for m in numeratorkids:
                        parentaai=[]
                        parentsheet=allkids[readablekids.index(m)][0]
                        for n in headingdict[parentsheet]:
                            if '_Area' in n:
                                parentaai.append(n)
                            elif 'Intensity' in n:
                                parentaai.append(n)
                        for n in headingdict.keys():
                            if HXM.middleofdelim(allkids[readablekids.index(m)][1]) == n:
                                childsheet=n
                        sharedaai=[]
                        for n in headingdict[childsheet]:
                            if n in parentaai:
                                sharedaai.append(n)
                        else:
                            valtofract=eg.multchoicebox(msg='Pick all the values you wish to do fraction analysis on for '+m,choices=sharedaai)
                            for r in valtofract:
                                fractanalysis.append(((parentsheet,r),(childsheet,r)))
                if len(fractanalysis)==0:
                    fractanalysis=False
                
            else:
                fractanalysis=False
            
            if eg.ynbox(msg='Do you want to compare the ratio of two wavelengths in any objects?'):
                waveratio=[]
                allwaves=[]
                readablewaves=[]
                for k in headingdict.keys():
                    for heads in headingdict[k]:
                        if 'Intensity' in heads:
                            allwaves.append((k,heads))
                            readablewaves.append(HXM.stripheaderonce(k)+' - '+HXM.stripheaderonce(heads))
                toberated=eg.multchoicebox(msg='Pick all the objects TO BE normalized', choices=readablewaves)
                if len(toberated)==0:
                    waveratio=False
                else:
                    for sinwave in toberated:
                        subreadablewaves=[]
                        sinwavepage=allwaves[readablewaves.index(sinwave)][0]
                        for fullwave in range(len(allwaves)):
                            if allwaves[fullwave][0]==sinwavepage:
                                subreadablewaves.append(readablewaves[fullwave])
                        raters=eg.multchoicebox(msg='Pick all the objects to normalize '+sinwave+' BY',choices=subreadablewaves)
                        if len(raters)==0:
                            pass
                        else:
                            for eachrate in raters:
                                waveratio.append((allwaves[readablewaves.index(sinwave)],allwaves[readablewaves.index(eachrate)]))
                
                    if len(waveratio)==0:
                        waveratio=False
                
            else:
                waveratio=False
                            
            if eg.ynbox(msg='Do you wish to save these settings as a default?'):
                defname=eg.enterbox('Enter a descriptive name for this default', strip=False)
                prevdef[defname]=tocompinit,dist,kidanalysis,normanalysis,fractanalysis,waveratio
                
        prevdef.close() 
    else:
        tocompinit, dist, kidanalysis, normanalysis,fractanalysis,waveratio=default

    #run analysis of distances        
    if tocompinit!=False:
        tocomp=[]
        if type(tocompinit)==list: #compatibility for old defaults
            for i in tocompinit:
                tocomp.append(parlist[strippedparlist.index(i)])
            pairdata={}
            for obj in tocomp:
                pairdata[obj]=pullallvals(rb,obj,["Location_Center_X","Location_Center_Y"],'Parent_Nuclei')
            pairslist=[]
            for ind1 in tocomp:
                templist=[]
                for ind2 in tocomp:
                    if ind1!=ind2:
                        templist.append((ind1,ind2))
                pairslist.append(templist)
            for startind in pairslist:
                workingsheetnum=sheetlist.index(tocompinit[tocomp.index(startind[0][0])])
                activesheetcols=numcoldict[workingsheetnum]
                wrsheet=wb.get_sheet(workingsheetnum)
                for pair in startind:
                    wrsheet.write(0,activesheetcols,str(tocompinit[tocomp.index(pair[0])])+'To'+str(tocompinit[tocomp.index(pair[1])]))
                    distance=compare2dicts(pairdata[pair[0]],pairdata[pair[1]],dist)
                    for i in distance.keys():
                        wrsheet.write(i,activesheetcols,distance[i])
                    numcoldict[workingsheetnum]+=1
                    activesheetcols+=1
        else:
            for i in tocompinit[0]:
                tocomp.append(parlist[strippedparlist.index(i)])
            pairdata={}
            for obj in tocomp:
                pairdata[obj]=pullallvals(rb,obj,["Location_Center_X","Location_Center_Y"],tocompinit[1])
            #print pairdata
            pairslist=[]
            for ind1 in tocomp:
                templist=[]
                for ind2 in tocomp:
                    if ind1!=ind2:
                        templist.append((ind1,ind2))
                pairslist.append(templist)
            for startind in pairslist:
                workingsheetnum=sheetlist.index(tocompinit[0][tocomp.index(startind[0][0])])
                activesheetcols=numcoldict[workingsheetnum]
                wrsheet=wb.get_sheet(workingsheetnum)
                for pair in startind:
                    wrsheet.write(0,activesheetcols,str(tocompinit[0][tocomp.index(pair[0])])+'To'+str(tocompinit[0][tocomp.index(pair[1])]))
                    distance=compare2dicts(pairdata[pair[0]],pairdata[pair[1]],dist)
                    for i in distance.keys():
                        wrsheet.write(i,activesheetcols,distance[i])
                    numcoldict[workingsheetnum]+=1
                    activesheetcols+=1
    
    if kidanalysis!=False:
        for i in kidanalysis:
            #print i, headingdict
            workingsheetnum=strippedparlist.index(HXM.stripheaderonce(i[0][0]))
            activesheetcols=numcoldict[workingsheetnum]
            wrsheet=wb.get_sheet(workingsheetnum)
            readsheet=rb.sheet_by_index(workingsheetnum)
            numerator=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index(i[0][1]))
            denominator=readsheet.col_values(headingdict[HXM.stripheaderonce(i[1][0])].index(i[1][1]))
            wrsheet.write(0,activesheetcols,'%'+HXM.middleofdelim(i[0][1])+'/'+HXM.middleofdelim(i[1][1]))
            for j in range(1,len(numerator)):
                if denominator[j]!=0:
                    wrsheet.write(j,activesheetcols,100*numerator[j]/denominator[j])
                else:
                    wrsheet.write(j,activesheetcols,0)
            numcoldict[workingsheetnum]+=1
           
          
    if normanalysis!=False:
        for i in normanalysis:
            workingsheetnum=strippedparlist.index(HXM.stripheaderonce(i[0][0]))
            activesheetcols=numcoldict[workingsheetnum]
            readsheet=rb.sheet_by_index(workingsheetnum)
            wrsheet=wb.get_sheet(workingsheetnum)
            images=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index('ImageNumber'))
            nucs=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index('Parent_'+i[1][0]))
            values=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index(i[0][1]))
            if i[2]=='Subtraction':
                subfact=makenormfactorsub(rb,parlist[strippedparlist.index(HXM.stripheaderonce(i[1][0]))],i[1][1])
                wrsheet.write(0,activesheetcols,HXM.stripheaderonce(values[0])+' minus '+subfact['name'])
                for j in range(1,readsheet.nrows):
                    wrsheet.write(j,activesheetcols,values[j]-subfact[(images[j],nucs[j])])
            else:
                divfact=makenormfactordiv(rb,parlist[strippedparlist.index(HXM.stripheaderonce(i[1][0]))],i[1][1])
                wrsheet.write(0,activesheetcols,HXM.stripheaderonce(values[0])+' divided by '+divfact['name'])
                for j in range(1,readsheet.nrows):
                    wrsheet.write(j,activesheetcols,values[j]/divfact[(images[j],nucs[j])])
            numcoldict[workingsheetnum]+=1
            
            
    if fractanalysis!=False:
        for i in fractanalysis:
            workingsheetnum=strippedparlist.index(HXM.stripheaderonce(i[0][0]))
            activesheetcols=numcoldict[workingsheetnum]
            wrsheet=wb.get_sheet(workingsheetnum)
            readsheet=rb.sheet_by_index(workingsheetnum)
            denominator=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index(i[0][1]))
            images=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index('ImageNumber'))
            nucs=readsheet.col_values(headingdict[HXM.stripheaderonce(i[0][0])].index('ObjectNumber'))
            nomdic=maketotal(rb,parlist[strippedparlist.index(HXM.stripheaderonce(i[1][0]))],i[1][1],i[0][0])
            wrsheet.write(0,activesheetcols,'% of '+HXM.stripheaderonce(denominator[0])+' of '+HXM.stripheaderonce(i[1][0])+'/'+HXM.stripheaderonce(i[0][0]))
            for j in range(1,len(denominator)):
                if (images[j],nucs[j]) in nomdic.keys():
                    numerator=nomdic[(images[j],nucs[j])]
                else:
                    numerator=0
                if denominator[j]!=0:
                    wrsheet.write(j,activesheetcols,100*numerator/denominator[j])
                else:
                    wrsheet.write(j,activesheetcols,0)
            numcoldict[workingsheetnum]+=1        
            
    if waveratio!=False:
        for i in waveratio:
            workingsheetnum=strippedparlist.index(HXM.stripheaderonce(i[0][0]))
            activesheetcols=numcoldict[workingsheetnum]
            wrsheet=wb.get_sheet(workingsheetnum)
            readsheet=rb.sheet_by_index(workingsheetnum)
            wrsheet.write(0,activesheetcols,'Ratio of '+HXM.stripheaderonce(i[0][1])+' to '+HXM.stripheaderonce(i[1][1]))
            for j in range(1,readsheet.nrows):
                #print (j,headingdict[HXM.stripheaderonce(i[0][0])].index(i[0][1]))
                ratio=readsheet.cell(j,headingdict[HXM.stripheaderonce(i[0][0])].index(i[0][1])).value/readsheet.cell(j,headingdict[HXM.stripheaderonce(i[1][0])].index(i[1][1])).value
                wrsheet.write(j,activesheetcols,ratio)
            numcoldict[workingsheetnum]+=1
            
    if default==None:
        r=os.path.join(directory,csvoutname+'.xls') #Save the excel file
    else:
        r=os.path.join(os.path.split(directory)[0],csvoutname+'.xls') #Save the excel file
    wb.save(r)

def batchmode():
    masterdir=eg.diropenbox("Which is the main experiment folder that has all of your subfolders with output files inside? (You'll choose which subfolders to use in a minute)")
    alldir=[]
    for i in os.listdir(masterdir):
        subfold=os.path.join(masterdir,i)
        if os.path.isdir(subfold):
            alldir.append(subfold)
    touse=eg.multchoicebox('Which of these do you want to analyze?',choices=alldir)
    prevdef=shelve.open(os.path.join(os.curdir,'csvshelf'),writeback=True)
    pickdefault=eg.choicebox(msg='Pick the default you wish to use', choices=prevdef.keys())
    default=prevdef[pickdefault]
    for folder in touse:
        calcnearestneighbor(folder,os.path.split(folder)[1],default)

if __name__=='__main__':

    batchmode()
