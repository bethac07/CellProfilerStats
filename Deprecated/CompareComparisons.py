import xlrd
import xlwt
from xlutils.copy import copy
from HandyXLModules import *
import easygui as eg
import os
from os.path import join

def runheatmap(sinbatch=False, batch=False,direct='direct',directname='direct'):
    parameters=[]
    idents=[]
    masterdictlist=[]
    benchmarks={}
    medbenchmarks={}
    if sinbatch==False:
        singleorpair=eg.indexbox('Do you want to compare single treatments or pairwise comparisons?',choices=['Pairwise','Single'])
    if batch==True:
        singleorpair=1
        for i in os.listdir(direct):
            if '.xls' in i:
                if 'ZZZZ' not in i:
                    idents.append(i)
                    a=xlrd.open_workbook(join(direct,i))
                    msheets=readsheets1file(a)
                    if 'Statistics' in msheets:
                        pickasheet='Statistics'
                    else:
                        pickasheet=eg.choicebox('Which sheet do you want to use?',choices=msheets)
                    sheetnum=msheets.index(pickasheet)
                    aa=copycol(a,sheetnum,0)
                    d={}
                    listofemptys=[0]
                    for i in range(1,len(aa)-1):
                        if aa[i]=='':
                            listofemptys.append(i)
                    listofemptys.append('')

                    for i in range(len(listofemptys)-2):
                        d[str(aa[listofemptys[i]+1])]=copysheet(a,sheetnum,int(listofemptys[i]+1),0,int(listofemptys[i+1]))
                        if str(aa[listofemptys[i]+1]) not in parameters:
                            parameters.append(str(aa[listofemptys[i]+1]))
                    d[str(aa[listofemptys[len(listofemptys)-2]+1])]=copysheet(a,sheetnum,int(listofemptys[len(listofemptys)-2]+1))
                    if str(aa[listofemptys[len(listofemptys)-2]+1]) not in parameters:
                        parameters.append(str(aa[listofemptys[len(listofemptys)-2]+1]))
                    masterdictlist.append(d)   
    else:
        addanother=True
        while addanother==True:
            choosebook=eg.fileopenbox(msg='Choose input Excel file',default='*.xls')
            idit=eg.enterbox('What should the identifier for this be?')
            idents.append(idit)
            a=xlrd.open_workbook(choosebook)
            msheets=readsheets1file(a)
            if 'Statistics' in msheets:
                pickasheet='Statistics'
            else:
                pickasheet=eg.choicebox('Which sheet do you want to use?',choices=msheets)
            sheetnum=msheets.index(pickasheet)
            aa=copycol(a,sheetnum,0)
            d={}
            listofemptys=[0]
            for i in range(1,len(aa)-1):
                if aa[i]=='':
                    listofemptys.append(i)
            listofemptys.append('')

            if singleorpair==0:
                for i in range(len(listofemptys)-2):
                    d[str(aa[listofemptys[i]+1])]=copysheet(a,sheetnum,int(listofemptys[i]+2),0,int(listofemptys[i+1]))
                    if str(aa[listofemptys[i]+1]) not in parameters:
                        parameters.append(str(aa[listofemptys[i]+1]))
                d[str(aa[listofemptys[len(listofemptys)-2]+1])]=copysheet(a,sheetnum,int(listofemptys[len(listofemptys)-2]+2))
            else:
                for i in range(len(listofemptys)-2):
                    d[str(aa[listofemptys[i]+1])]=copysheet(a,sheetnum,int(listofemptys[i]+1),0,int(listofemptys[i+1]))
                    if str(aa[listofemptys[i]+1]) not in parameters:
                        parameters.append(str(aa[listofemptys[i]+1]))
                d[str(aa[listofemptys[len(listofemptys)-2]+1])]=copysheet(a,sheetnum,int(listofemptys[len(listofemptys)-2]+1))
            if str(aa[listofemptys[len(listofemptys)-2]+1]) not in parameters:
                parameters.append(str(aa[listofemptys[len(listofemptys)-2]+1]))
            masterdictlist.append(d)        
            again=eg.ynbox('Do you want to add another sheet or file?')
            if not again:
                addanother=False
    setctl=eg.choicebox('Which of these should be the baseline control?',choices=idents)
    ctlindex=idents.index(setctl)
    if batch==True:
        newfilename=join(direct,directname+'Heat.xls')
        writebook=xlwt.Workbook()
        wsheet=writebook.add_sheet(directname+'Heat')
    else:
        reuse=eg.indexbox('Where to save the output?',choices=['Create a new file', 'Make a new sheet in another file'])
        if reuse==0: #If the user wants to make a new file
            newfilename=eg.filesavebox(msg='What do you want to name the file?',filetypes=["*.xls"])+'.xls'
            writebook=xlwt.Workbook() #make a new file
            newsheetname=eg.enterbox('What do you want to call the new sheet?')
            wsheet=writebook.add_sheet(newsheetname) #give it a sheet
        if reuse==1: #If the user wants a new sheet in an old file
            openfilename=eg.fileopenbox(msg='Choose Excel file to write to',default='*.xls')
            otherbook=xlrd.open_workbook(openfilename) #Open it
            writebook=copy(otherbook) #make it writable
            newsheetname=eg.enterbox('What do you want to call the new sheet?')
            wsheet=writebook.add_sheet(newsheetname) #add the new sheet
    indivwidth=maxnumcols(masterdictlist[0].items()[1][1])
    totalwidth=len(idents)*indivwidth
    totalheight=max(len(masterdictlist[0].items()[0][1]),len(masterdictlist[0].items()[1][1]))
    stylecenter=xlwt.easyxf('alignment: horizontal center')
    stylevsmall = xlwt.easyxf('pattern: fore_colour sky_blue, pattern solid_fill')
    stylesmall = xlwt.easyxf('pattern: fore_colour pale_blue, pattern solid_fill')
    stylebig = xlwt.easyxf('pattern: fore_colour rose, pattern solid_fill')
    stylevbig = xlwt.easyxf('pattern: fore_colour red,pattern solid_fill')
    stylemsmall = xlwt.easyxf('pattern: fore_colour light_green, pattern solid')
    stylemvsmall = xlwt.easyxf('pattern: fore_colour bright_green,pattern solid_fill')
    stylembig = xlwt.easyxf('pattern: fore_colour lavender, pattern solid_fill')
    stylemvbig = xlwt.easyxf('pattern: fore_colour purple_ega, pattern solid_fill')
    for i in range(len(idents)):
        wsheet.write_merge(0,0,(indivwidth*i)+1,(indivwidth*(i+1)),idents[i],stylecenter)
    writerow=1
    for i in range(len(parameters)):
        if parameters[i] in masterdictlist[ctlindex]:
            f=[]
            for k in masterdictlist[ctlindex][parameters[i]]:
                if type(k[1])==float:
                    f.append(k[1])
            if singleorpair==0:
                if f!=[]:
                    f.sort()
                    ff=[f[0]-(abs(f[0]-f[1])),f[0]-(0.5*abs(f[0]-f[1])),f[1]+(0.5*abs(f[0]-f[1])),f[1]+(abs(f[0]-f[1]))]
            else:
                if f!=[]:
                    ff=[f[0]*0.33,f[0]*0.67,f[0]*1.5,f[0]*3]
            benchmarks[parameters[i]]=ff
            g=[]
            for k in masterdictlist[ctlindex][parameters[i]]:
                if type(k[4])==float:
                    g.append(k[4])
            if singleorpair==0:
                if g!=[]:
                    g.sort()
                    gg=[g[0]-(abs(g[0]-g[1])),g[0]-(0.5*abs(g[0]-g[1])),g[1]+(0.5*abs(g[0]-g[1])),g[1]+(abs(g[0]-g[1]))]
            else:
                if g!=[]:
                    gg=[g[0]*0.33,g[0]*0.67,g[0]*1.5,g[0]*3]
            medbenchmarks[parameters[i]]=gg
        else:
            benchmarks[parameters[i]]=''
            medbenchmarks[parameters[i]]=''
    for i in range(len(parameters)):
        wsheet.write_merge(writerow,writerow,1,totalwidth+1,str(parameters[i]),stylecenter)
        startcol=1
        for j in range(len(masterdictlist)):
            if parameters[i] in masterdictlist[j]:
                if benchmarks[parameters[i]]=='':
                     writesheet(wsheet,masterdictlist[j][parameters[i]],writerow+1,startcol)
                else:
                    for row in range(len(masterdictlist[j][parameters[i]])):
                        for col in ([0,2,3]+range(5,indivwidth)):
                            wsheet.write(row+writerow+1,col+startcol,masterdictlist[j][parameters[i]][row][col])
                    for row in range(len(masterdictlist[j][parameters[i]])):
                        if type(masterdictlist[j][parameters[i]][row][1])==float:
                            num=masterdictlist[j][parameters[i]][row][1]
                            bench=benchmarks[parameters[i]]
                            if num<bench[0]:
                                wsheet.write(row+writerow+1,1+startcol,masterdictlist[j][parameters[i]][row][1],stylevsmall)
                            elif num>=bench[0] and num<bench[1]:
                                wsheet.write(row+writerow+1,1+startcol,masterdictlist[j][parameters[i]][row][1],stylesmall)
                            elif num>bench[2] and num<=bench[3]:
                                wsheet.write(row+writerow+1,1+startcol,masterdictlist[j][parameters[i]][row][1],stylebig)
                            elif num>bench[3]:
                                wsheet.write(row+writerow+1,1+startcol,masterdictlist[j][parameters[i]][row][1],stylevbig)
                            else:
                                wsheet.write(row+writerow+1,1+startcol,masterdictlist[j][parameters[i]][row][1])
                        else:
                            wsheet.write(row+writerow+1,1+startcol,masterdictlist[j][parameters[i]][row][1])
                        if type(masterdictlist[j][parameters[i]][row][4])==float:
                            num=masterdictlist[j][parameters[i]][row][4]
                            medbench=medbenchmarks[parameters[i]]
                            if num<medbench[0]:
                                wsheet.write(row+writerow+1,4+startcol,masterdictlist[j][parameters[i]][row][4],stylemvsmall)
                            elif num>=medbench[0] and num<medbench[1]:
                                wsheet.write(row+writerow+1,4+startcol,masterdictlist[j][parameters[i]][row][4],stylemsmall)
                            elif num>medbench[2] and num<=medbench[3]:
                                wsheet.write(row+writerow+1,4+startcol,masterdictlist[j][parameters[i]][row][4],stylembig)
                            elif num>medbench[3]:
                                wsheet.write(row+writerow+1,4+startcol,masterdictlist[j][parameters[i]][row][4],stylemvbig)
                            else:
                                wsheet.write(row+writerow+1,4+startcol,masterdictlist[j][parameters[i]][row][4])
                        else:
                            wsheet.write(row+writerow+1,4+startcol,masterdictlist[j][parameters[i]][row][4])
            startcol+=indivwidth
        writerow+=(totalheight+1)
    if batch==True:
        writebook.save(newfilename)
    else:
        if reuse==0: #if we're saving a new file, save it with the name the user input
            writebook.save(newfilename)
        else:
            writebook.save(openfilename)


def runbatchmode(direct):
    for i in os.listdir(direct):
        if '.xls' not in i:
            for j in os.listdir(join(direct,i)):
                if '.xls' not in j:
                    runheatmap(True,True,join(direct,i,j),j)
                
"""

"""
