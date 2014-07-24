# -*- coding: utf-8 -*-
"""
Created on Sat Oct 29 13:10:37 2011

@author: Beth Cimini
"""

from scipy import stats
import numpy
import xlrd

from HandyXLModules import *

def man2histogram(book,sheetindex):
    headings=colheadingreadernum(book,sheetindex)
    TIF1headings=[]
    TIF2headings=[]
    TIF1data=[]
    TIF2data=[]
    for i in range(len(headings)):
        if i%2==0:
            TIF1headings.append(headings[i])
            TIF1data.append(copycol(book, sheetindex,i,1))
        else:
            TIF2headings.append(headings[i])
            TIF2data.append(copycol(book, sheetindex,i,1))
    
    for i in TIF1data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)
    
    for i in range(len(TIF1data)):
        median=numpy.median(TIF1data[i])
        if i>0:
            ksvalue,pvalue=stats.ks_2samp(TIF1data[0],TIF1data[i])
            TIF1headings[i]+=' median= '+"%.2f" %median+' p='+"%0.2f" %pvalue
        else:
            TIF1headings[i]+=' median= '+"%.2f" %median        
    
    for i in TIF2data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)
    
    for i in range(len(TIF2data)):
        median=numpy.median(TIF2data[i])
        if i>0:
            ksvalue,pvalue=stats.ks_2samp(TIF2data[0],TIF2data[i])
            TIF2headings[i]+=' median= '+"%.2f" %median+' p='+"%0.2f" %pvalue
        else:
            TIF2headings[i]+=' median= '+"%.2f" %median
    
    graphmanyhists(TIF1data, TIF1headings, '% of TRF1 telomeres that are TIFs',r'D:\20110810\2DStills\PercentTIF1s_TIN2Smed')
    graphmanyhists(TIF2data, TIF2headings, '% of TIN2 telomeres that are TIFs',r'D:\20110810\2DStills\PercentTIFTIN2s_TIN2Smed')
    
def manhistogram(book,sheetindex):
    headings=colheadingreadernum(book,sheetindex)

    TIF1data=[]
    for i in range(len(headings)):
        TIF1data.append(copycol(book, sheetindex,i,1))

    for i in TIF1data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)
    
    for i in range(len(TIF1data)):
        median=numpy.median(TIF1data[i])
        if i>0:
            ksvalue,pvalue=stats.ks_2samp(TIF1data[0],TIF1data[i])
            headings[i]+=' median= '+"%.2f" %median+' p='+"%0.2f" %pvalue
        else:
            headings[i]+=' median= '+"%.2f" %median

    graphmanyhists(TIF1data, headings, '% of TRF1 telomeres that are TIFs',r'D:\20110810\2DStills\PercentTIF1s_noTIN2S')

def manscatter(book,sheetindex):
    headings=colheadingreadernum(book,sheetindex)
    Factheadings=[]
    Fact1data=[]
    Fact2data=[]
    for i in range(len(headings)):
        if i%2==0:
            Factheadings.append(headings[i])
            Fact1data.append(copycol(book, sheetindex,i,1))
        else:
            Fact2data.append(copycol(book, sheetindex,i,1))
    
    for i in Fact1data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)

    for i in Fact2data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)

    graphscatter(Fact1data,Fact2data,Factheadings,r'F:\20111010\20111010NucleiScatter','TRF1 Integrated Intensity', 'TRF2 Integrated Intensity',mantitle='Nuclei-TRF1 vs TRF2 Integrated Intensity', savefiles=True )

def manbubble(book,sheetindex):
    headings=colheadingreadernum(book,sheetindex)
    Factheadings=[]
    Fact1data=[]
    Fact2data=[]
    Fact3data=[]
    for i in range(len(headings)):
        if i%3==0:
            Factheadings.append(headings[i])
            Fact3data.append(copycol(book, sheetindex,i,1))
        elif i%3==1:
            Fact1data.append(copycol(book, sheetindex,i,1))
        else:
            Fact2data.append(copycol(book, sheetindex,i,1))
    
    for i in Fact1data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)

    for i in Fact2data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)
                
    for i in Fact3data:
        for j in range(len(i)-1,0,-1):
            if type(i[j])!=float:
                i.pop(j)

    graphbubble(Fact1data,Fact2data,Fact3data,Factheadings,r'F:\20110729\ActinFociScatter','Actin Channel Integrated Intensity','Telomere Channel Integrated Intensity','DNA Damage Channel Integrated Intensity', mantitle='Actin Foci-Integrated Intensity of Actin Channel vs Telomere Channel vs DNA Damage Channel' ,savefiles=True)

a=xlrd.open_workbook(r'D:\20110810\2DStills\PercentTIFs.xls')
manhistogram(a,0)
man2histogram(a,1)

