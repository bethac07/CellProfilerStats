# -*- coding: utf-8 -*-
"""
Created on Sat Sep 21 20:31:09 2013

@author: Beth Cimini
"""
import xlrd
import matplotlib as mpl
from mpl_toolkits.mplot3d import Axes3D, art3d
from matplotlib._png import read_png
from pylab import ogrid
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.cbook import get_sample_data

a=xlrd.open_workbook(r'G:\TrackingOutputFolders200Frames\RPE-CRISPR50frameMin.xls')
sheet=a.sheet_by_index(1)
spotindices=[]
for m in range(59679,sheet.nrows):
    if sheet.cell(m,0).value!='':
        spotindices.append(m)
#print spotindices
allx=[]
ally=[]
allz=[]
for starts in range(len(spotindices)-1):
    begin=sheet.cell(spotindices[starts],1).value%200
    thisx=[]
    thisy=[]
    thisz=[]
    for frames in range(spotindices[starts]+1,spotindices[starts+1]):
        if sheet.cell(frames,1).value+begin<=50:
            thisx.append(sheet.cell(frames,2).value)
            thisy.append(sheet.cell(frames,3).value)
            thisz.append(sheet.cell(frames,1).value+begin)
    print thisx, thisy, thisz
    allx.append(thisx)
    ally.append(thisy)
    allz.append(thisz)

plt.ioff()
fig = plt.figure()
#ay=fig.add_subplot(2,1,1)
#fn = get_sample_data("lena.png", asfileobj=False)
#rawim=read_png(fn)
ax = fig.gca(projection='3d')
rawim=plt.imread(r'G:\TrackingOutputFolders200Frames\RPE-CRISPR\movie_0018_488-1aligned1_200_1_InputImage.png')
#ax.imshow(rawim,cmap='gray',aspect='equal')
#ax=fig.add_subplot(2,1,1,projection='3d')

#rawim=read_png(r'G:\Telomeredynamics_RPE_V3_SC_80ul\200framesOutPng\movie_0001_488-1aligned1_200_1_InputImage.png')
#print rawim
x,y= ogrid[0:rawim.shape[0],0:rawim.shape[1]]
ax.plot_surface(x,y,0,rstride=5,cstride=5,facecolors=rawim,cmap='gray')
for i in range(len(allx)):
    ax.plot(ally[i], allx[i],allz[i])
#plt.axis('image')
ax.view_init(elev=53, azim=16)
#imgplot = plt.imshow(rawim)
plt.show()       
        
        

"""import matplotlib as mpl
from mpl_toolkits.mplot3d import Axes3D, art3d
from pylab import ogrid
import matplotlib.pyplot as plt

plt.ioff()
fig = plt.figure()
ay=fig.add_subplot(2,1,1)
rawim=plt.imread(r'G:\Path\myimage.png')
ay.imshow(rawim,cmap='gray')
ax=fig.add_subplot(2,1,2,projection='3d')
x,y= ogrid[0:rawim.shape[0],0:rawim.shape[1]]
ax.plot_surface(x,y,0,rstride=5,cstride=5,facecolors=rawim,cmap='gray')
ax.view_init(elev=45, azim=12)
plt.show()  """    
        
    