import xlrd
from scipy import stats
from HandyXLModules import *
from CompParamSingle import arrangedivsin
import datetime
import easygui as eg



def choosewhichstatssin(book,book2):
    map2={}
    t=[]
    colsbysheet=[]
    a=readsheets1file(book)
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
    if eg.ynbox(msg='Do you want to use the defaults?'):
        g=[[[0, 2, 1, 0], [0, 3, 1, 0], [0, 8, 1, 0], [0, 9, 1, 0], [0, 16, 1, 1,
            ['<', '500']], [0, 18, 1, 0], [0, 18, 1, 1, ['>', '0']], [0, 19, 1, 0],
            [0, 19, 1, 1, ['>', '0']], [4, 2, 1, 0], [5, 2, 1, 0], [6, 2, 1, 0],
            [6, 3, 1, 0], [6, 8, 1, 0], [6, 9, 1, 0], [6, 16, 1, 1, ['<', '500']],
            [6, 18, 1, 0], [6, 18, 1, 1, ['>', '0']], [6, 19, 1, 0], [6, 19, 1, 1,
            ['>', '0']], [8, 2, 1, 0], [8, 3, 1, 0], [8,4, 1, 0], [8, 5, 1, 0],
            [8, 6, 1, 0], [8, 7, 1, 0], [8, 8, 1, 0], [12, 2, 1, 0], [13, 2, 1, 0],
            [14, 2, 1, 0], [14, 3, 1, 0], [14, 8, 1, 0], [14, 9, 1, 0], [14, 16, 1, 1,
            ['<', '500']], [14, 18, 1, 0], [14, 18, 1, 1, ['>', '0']], [14, 19, 1, 0],
            [14, 19, 1, 1, ['>', '0']]], [[0, 2, 1, 0], [0, 3, 1, 0], [0, 8, 1, 0],
            [0, 9, 1, 0], [0, 16, 1, 1, ['<', '500']], [0, 18, 1, 0], [0, 18, 1, 1,
            ['>', '0']], [0, 19, 1, 0], [0, 19, 1, 1, ['>', '0']], [4, 2, 1, 0],
            [5, 2, 1, 0], [6, 2,1, 0], [6, 3, 1, 0], [6, 8, 1, 0], [6, 9, 1, 0],
            [6, 16, 1, 1, ['<', '500']], [6, 18, 1, 0], [6, 18, 1, 1, ['>', '0']],
            [6, 19, 1, 0], [6, 19, 1, 1, ['>', '0']], [8, 2, 1, 0], [8, 3, 1, 0],
            [8, 4, 1, 0], [8, 5, 1, 0], [8, 6, 1, 0], [8, 7, 1, 0], [8, 8, 1, 0],
            [12, 2, 1, 0], [13, 2, 1, 0], [14, 2, 1, 0], [14, 3, 1, 0],[14, 8, 1, 0],
            [14, 9, 1, 0], [14, 16, 1, 1, ['<', '500']], [14, 18, 1, 0],
            [14, 18, 1, 1, ['>', '0']], [14, 19, 1, 0], [14, 19, 1, 1, ['>', '0']]]]

    else:
        #Pull all the possible parameters
        
        xparams=[]
        fx=eg.multchoicebox(msg='Select parameters',choices=t)
        graphx=fx
        filtf=eg.multchoicebox(msg='Do you want to filter any of these?',choices=(fx))
        for param in xrange(len(fx)):
            s=[]
            s.append(int(fx[param][0:2])) #first item is the sheet index
            s.append(int(fx[param][4:7])) #second item is the column index
            if fx[param] in graphx:
                s.append(1)#pass the graphing option
            else:
                s.append(0)

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

            xparams.append(s)
        
        xparams2=[]
        for i in xrange(len(xparams)):
            mapped=map2[(xparams[i][0],xparams[i][1])]
            xparams2.append([mapped[0],mapped[1]]+xparams[i][2:])
            
        g=[xparams,xparams2] #return the list of parameters the user wants to compare
        print g
    return g

def findtheps(list1,list2):
    p1=[]
    p2=[]
    p3=[]
    p4=[]
    p5=[]
    p6=[]
    p7=[]
    p8=[]
    for i in range(len(list1)):
        uvalue,pvalue=stats.mannwhitneyu(list1[i][1:],list2[i][1:])
        ksvalue,pvalue2=stats.ks_2samp(list1[i][1:],list2[i][1:])
        if pvalue <=0.1:
            if pvalue<=0.05:
                if pvalue<=0.01:
                    if pvalue<=0.001:
                        p4.append((list1[i][0][0],pvalue))
                    else:
                        p3.append((list1[i][0][0],pvalue))
                else:
                    p2.append((list1[i][0][0],pvalue))
            else:
                p1.append((list1[i][0][0],pvalue))
        if pvalue2 <=0.1:
            if pvalue2<=0.05:
                if pvalue2<=0.01:
                    if pvalue2<=0.001:
                        p8.append((list1[i][0][0],pvalue2))
                    else:
                        p7.append((list1[i][0][0],pvalue2))
                else:
                    p6.append((list1[i][0][0],pvalue2))
            else:
                p5.append((list1[i][0][0],pvalue2))
    ps=[p1,p2,p3,p4,p5,p6,p7,p8]
    return ps

def writethestuff(filename,listofps):
    filename.write('Mann Whitney p<0.1: \n')
    filename.write(str(listofps[0])+'\n\n')
    filename.write('Mann Whitney p<0.05: \n')
    filename.write(str(listofps[1])+'\n\n')
    filename.write('Mann Whitney p<0.01: \n')
    filename.write(str(listofps[2])+'\n\n')
    filename.write('Mann Whitney p<0.001: \n')
    filename.write(str(listofps[3])+'\n\n')
    filename.write('KS p<0.1: \n')
    filename.write(str(listofps[4])+'\n\n')
    filename.write('KS p<0.05: \n')
    filename.write(str(listofps[5])+'\n\n')
    filename.write('KS p<0.01: \n')
    filename.write(str(listofps[6])+'\n\n')
    filename.write('KS p<0.001: \n')
    filename.write(str(listofps[7])+'\n\n\n\n')

def dothestuff():
    direct=eg.diropenbox()
    ident=eg.enterbox('What identifier do you want to give this experiment?')
    a=findexcel(direct)
    b=[]
    c=[]
    for i in a:
        b.append(i[0])
        c.append(i[1])
    whichfiles=eg.multchoicebox(msg='Which files do you want to use?', choices=b)
    baseline=eg.choicebox(msg='Which file is the baseline?', choices=whichfiles)
    whichfiles.remove(baseline)
    basebook=xlrd.open_workbook(c[b.index(baseline)])
    runeach=eg.ynbox('Do you want to run the same statistics for all files, and are they all identically laid out?')
    saveto=eg.filesavebox('Where do you want to save the information?')
    if '.txt' not in saveto:
        saveto+= '.txt'
    writeto=open(saveto,'a')
    writeto.write('For the experiment '+ident+', the following things are significant as compared to '+baseline[0:len(baseline)-4]+':\n\n')               
    if runeach:
        compbook=xlrd.open_workbook(c[b.index(whichfiles[0])])
        whichstats=choosewhichstatssin(basebook,compbook)
        basevals=arrangedivsin(basebook,whichstats[0])
        for i in whichfiles:
            writeto.write(i+': \n\n')
            compbook=xlrd.open_workbook(c[b.index(i)])
            compvals=arrangedivsin(compbook,whichstats[1])
            ps=findtheps(basevals,compvals)
            writethestuff(writeto,ps)
    else:
        for i in whichfiles:
            writeto.write(i+': \n')
            compbook=xlrd.open_workbook(c[b.index(i)])
            whichstats=choosewhichstatssin(basebook,compbook)
            basevals=arrangedivsin(basebook,whichstats[0])
            compvals=arrangedivsin(compbook,whichstats[1])
            ps=findtheps(basevals,compvals)
            writethestuff(writeto,ps)

    writeto.write('Generated on '+datetime.datetime.ctime(datetime.datetime.today())+'\n\n\n\n')
    writeto.close()

if __name__=='__main__':
    dothestuff()

"""
Default for old analysis
[[[0, 2, 1, 0], [0, 3, 1, 0], [0, 8, 1, 0], [0, 9, 1, 0], [0, 16, 1, 1, ['<', '500']], [0, 18, 1, 0],
           [0, 18, 1, 1, ['>', '0']], [0, 19, 1, 0], [0, 19, 1, 1, ['>', '0']], [4, 2, 1, 0], [5, 2, 1, 0],
           [6, 2, 1, 0], [6, 3, 1, 0], [6, 8, 1, 0],[6, 9, 1, 0], [6, 16, 1, 1, ['<', '500']], [6, 18, 1, 0],
           [6, 18, 1, 1, ['>', '0']], [6, 19, 1, 0], [6, 19, 1, 1, ['>', '0']], [8, 2, 1, 0], [8, 3, 1, 0],
           [8, 4, 1, 0], [8, 5, 1, 0], [8, 6, 1, 0], [8, 7, 1, 0], [8, 8, 1, 0], [10, 2, 1, 0],[11, 2, 1, 0],
           [12, 2, 1, 0], [12, 3, 1, 0], [12, 8, 1, 0], [12, 9, 1, 0], [12,16, 1, 1, ['<', '500']],
           [12, 18, 1, 0], [12, 18, 1, 1, ['>', '0']], [12, 19, 1, 0], [12, 19, 1, 1, ['>', '0']]],
           [[0, 2, 1, 0], [0, 3, 1, 0], [0, 8, 1, 0], [0, 9, 1, 0], [0, 16, 1, 1, ['<', '500']], [0, 18, 1, 0],
           [0, 18, 1, 1, ['>', '0']], [0, 19, 1, 0], [0, 19, 1, 1, ['>', '0']], [4, 2, 1, 0], [5, 2, 1, 0],
           [6, 2, 1, 0], [6, 3, 1, 0], [6, 8, 1, 0],[6, 9, 1, 0], [6, 16, 1, 1, ['<', '500']], [6, 18, 1, 0],
           [6, 18, 1, 1, ['>', '0']], [6, 19, 1, 0], [6, 19, 1, 1, ['>', '0']], [8, 2, 1, 0], [8, 3, 1, 0],
           [8, 4, 1, 0], [8, 5, 1, 0], [8, 6, 1, 0], [8, 7, 1, 0], [8, 8, 1, 0], [10, 2, 1, 0],[11, 2, 1, 0],
           [12, 2, 1, 0], [12, 3, 1, 0], [12, 8, 1, 0], [12, 9, 1, 0], [12,16, 1, 1, ['<', '500']],
           [12, 18, 1, 0], [12, 18, 1, 1, ['>', '0']], [12, 19, 1, 0], [12, 19, 1, 1, ['>', '0']]]]

Default for big foci
[[[0, 2, 1, 0], [0, 3, 1, 0], [0, 8, 1, 0], [0, 9, 1, 0], [0, 16, 1, 1,
['<', '500']], [0, 18, 1, 0], [0, 18, 1, 1, ['>', '0']], [0, 19, 1, 0],
[0, 19, 1, 1, ['>', '0']], [4, 2, 1, 0], [5, 2, 1, 0], [6, 2, 1, 0],
[6, 3, 1, 0], [6, 8, 1, 0], [6, 9, 1, 0], [6, 16, 1, 1, ['<', '500']],
[6, 18, 1, 0], [6, 18, 1, 1, ['>', '0']], [6, 19, 1, 0], [6, 19, 1, 1,
['>', '0']], [8, 2, 1, 0], [8, 3, 1, 0], [8,4, 1, 0], [8, 5, 1, 0],
[8, 6, 1, 0], [8, 7, 1, 0], [8, 8, 1, 0], [12, 2, 1, 0], [13, 2, 1, 0],
[14, 2, 1, 0], [14, 3, 1, 0], [14, 8, 1, 0], [14, 9, 1, 0], [14, 16, 1, 1,
['<', '500']], [14, 18, 1, 0], [14, 18, 1, 1, ['>', '0']], [14, 19, 1, 0],
[14, 19, 1, 1, ['>', '0']]], [[0, 2, 1, 0], [0, 3, 1, 0], [0, 8, 1, 0],
[0, 9, 1, 0], [0, 16, 1, 1, ['<', '500']], [0, 18, 1, 0], [0, 18, 1, 1,
['>', '0']], [0, 19, 1, 0], [0, 19, 1, 1, ['>', '0']], [4, 2, 1, 0],
[5, 2, 1, 0], [6, 2,1, 0], [6, 3, 1, 0], [6, 8, 1, 0], [6, 9, 1, 0],
[6, 16, 1, 1, ['<', '500']], [6, 18, 1, 0], [6, 18, 1, 1, ['>', '0']],
[6, 19, 1, 0], [6, 19, 1, 1, ['>', '0']], [8, 2, 1, 0], [8, 3, 1, 0],
[8, 4, 1, 0], [8, 5, 1, 0], [8, 6, 1, 0], [8, 7, 1, 0], [8, 8, 1, 0],
[12, 2, 1, 0], [13, 2, 1, 0], [14, 2, 1, 0], [14, 3, 1, 0],[14, 8, 1, 0],
[14, 9, 1, 0], [14, 16, 1, 1, ['<', '500']], [14, 18, 1, 0],
[14, 18, 1, 1, ['>', '0']], [14, 19, 1, 0], [14, 19, 1, 1, ['>', '0']]]]"""
