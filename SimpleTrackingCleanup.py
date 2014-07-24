# -*- coding: utf-8 -*-
"""
Created on Tue Mar 01 16:41:44 2011

@author: Beth Cimini
"""
#from PIL import Image
import xlrd
import xlwt
import xlutils
from xlutils.copy import copy
import easygui as eg
import HandyXLModules as HXM
import Spots
import numpy
from scipy import stats
import os
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib._png import read_png
from pylab import ogrid
import PIL

def lastunderscoresuffix(str):
    undind=0
    for i in range(len(str)):
        if str[i]=='_':
            undind=i
    return str[undind:]

def imagefilenames(filename):
    for i in os.listdir(os.path.split(filename)[0]):
            if '_Image.csv' in i:
                array3=numpy.genfromtxt(os.path.join(os.path.split(filename)[0],i),delimiter=',',dtype=None)
    
    diffiles=[]    
    difpaths=[]
    imageheadings=list(array3[0,:])
    #print imageheadings
    firstrow=list(array3[1,:])
    for i in range(len(imageheadings)):
        if 'FileName' in imageheadings[i]:
            diffiles.append(i)
        if 'Path' in imageheadings[i]:
            difpaths.append(i)
    if len(diffiles)>1:
        checkfirsts=[]
        for i in diffiles:
            if firstrow[i] not in checkfirsts:
                checkfirsts.append(firstrow[i])
        if len(checkfirsts)>1:
            filenameuse=eg.choicebox('Which of these is the image you want to draw on?',choices=diffiles)
            files=list(array3[:,imageheadings.index(filenameuse)])
            path=list(array3[:,imageheadings.index('Path'+filenameuse[4:])])
        else:
            files=list(array3[:,diffiles[0]])
            path=list(array3[:,imageheadings.index('Path'+imageheadings[i][4:])])    
    else:
        files=list(array3[:,diffiles[0]])
        path=list(array3[:,difpaths[0]])
            
    fullfilename={}
    firstfilename={}
    count=0
    imcount=1
    for i in range(1,len(path)):
        if files[i]!=files[i-1]:
            count=0
            firstfilename[imcount]=os.path.join(path[i],files[i])
            imcount+=1
        else:
            count+=1
        fullfilename[i]=(os.path.join(path[i],files[i]),count)
        
    #print fullfilename
    return fullfilename,firstfilename
    
        
def runtrackingcleanup(filein,minlen,maxlen,identifier,addonlist,fileout,makemovies,todrawsuffix,movielen,framerate,pixelsize):
    
    otherbook=xlrd.open_workbook(fileout) #Open it
    writebook=copy(otherbook) #make it writable
    writesheet=writebook.add_sheet(identifier) #add the new sheet

    e=Spots.spots(filein,movielen,minlen,maxlen) #initialize the spot instances
    #print e
    print 'made spots'
    f=e.xyandrealt()
    #print 'xyandrealt'
    spotheadings=e[0]
    
    #print names
    """plt.ioff()
fig = plt.figure()
rawim=read_png(r'G:\testmage.png')
x,y= ogrid[0:rawim.shape[0],0:rawim.shape[1]]
ax = fig.gca(projection='3d')
ax.plot_surface(x,y,0.0,rstride=2,cstride=2,facecolors=rawim,color='w')
for i in range(len(allx)):
    ax.plot(ally[i], allx[i], allz[i])
ax.view_init(elev=65, azim=12)"""
    
    if makemovies==True:
        #print 'movies'
    
        #colorlist=['black','blue','crimson','darkorange','darkorchid','darkcyan','burlywood','deeppink','lawngreen','indigo','seashell','orangered','olive','powderblue','yellow']
        colorlist=['r','c','m','y','k','b','g']
        names,percellnames=imagefilenames(filein)
        tracks,tracksfor3d=e.trackstonow(movielen)
        """for i in tracksfor3d.keys():
            #print percellnames[i]
            plt.ioff()
            fig = plt.figure()
            imagein=os.path.join(os.path.split(filein)[0],os.path.split(percellnames[i])[1][:os.path.split(percellnames[i])[1].index('.')]+'_0'+todrawsuffix)
            #print percellnames[i],imagein
            rawim=plt.imread(imagein)
            #print rawim
            x,y= ogrid[0:rawim.shape[0],0:rawim.shape[1]]
            ax = fig.gca(projection='3d')
            #ax.plot_surface(x,y,0.0,rstride=1,cstride=1,facecolors=rawim,color='w')
            for j in tracksfor3d[i]:
                ax.plot(j[1], j[0], j[2])
            ax.view_init(elev=50, azim=12)
            istring='%02d' %i
            plt.savefig(os.path.join(os.path.split(fileout)[0],identifier+'Cell'+istring+'_3DTracks.png'))
            plt.close() """   
    
    
        for i in tracks.keys():
            subfolder=os.path.split(names[i][0])[1]
            subfoldername=subfolder[:subfolder.index('.')]+os.path.split(filein)[1][os.path.split(filein)[1].index('_'):os.path.split(filein)[1].index('.')]+'_tracks'       
            newdir=os.path.join(os.path.split(filein)[0],subfoldername)
            if not os.path.isdir(newdir):
                os.mkdir(newdir)
                #c=names[i][:names[i].index('.')]+'_'+str(i)+'_raw.bmp'
            dotindex=0
            for eachchar in range(len(os.path.split(names[i][0])[1])):
                if os.path.split(names[i][0])[1][eachchar]=='.':
                    dotindex=eachchar
            imagein=os.path.join(os.path.split(filein)[0],os.path.split(names[i][0])[1][:dotindex]+str(names[i][1])+todrawsuffix)
            plt.ioff()
            fig = plt.figure()
            rawim=plt.imread(imagein)
            implot=plt.imshow(rawim,cmap='gray')
            plt.axis('image')
            plt.axis('off')
            plt.savefig(os.path.join(newdir,subfolder[:subfolder.index('.')]+'_'+str(names[i][1])+'_rawmovies.png'),bbox_inches='tight',pad_inches=0)
            for eachtrack in tracks[i]:
                #plt.plot(eachtrack[0],eachtrack[1])
                plt.plot(eachtrack[0],eachtrack[1],marker='o',markersize=7,fillstyle='none',color=colorlist[eachtrack[2]])
            plt.axis('image')
            plt.axis('off')
            plt.savefig(os.path.join(newdir,subfolder[:subfolder.index('.')]+'_'+str(names[i][1])+'movies.png'),bbox_inches='tight',pad_inches=0)
            plt.close()
        
            
    #print 'finished drawing'
    
    col=4
    stufftowrite=[]
    meanchannels=[]
    for i in spotheadings:
        if 'MeanIntensity' in i:
            meanchannels.append(i[(i.index('MeanIntensity')+14):])
        if 'IntegratedIntensity' in i:
            channelname=i[(i.index('IntegratedIntensity')+20):]
            intoutput,medint=e.intint(channelname)
            tracksbycell=e.intintandintdistpercell(movielen,channelname,pixelsize)
            addonlist[7][identifier]=tracksbycell
            if identifier in addonlist[0].keys():
                addonlist[0][identifier].append(medint+[channelname])
            else:
                addonlist[0][identifier]=[medint+[channelname]]
            """intfilt=[]
            for j in intoutput:
                intfilt.append(j[1:])"""
            stufftowrite.append(intoutput)
            writesheet.write(0,col, i)
            col+=1
        if 'AreaShape_Area' in i:
            sizeoutput,medsize=e.size(pixelsize)
            addonlist[1][identifier]=medsize
            """sizefilt=[]
            for j in sizeoutput:
                sizefilt.append(j[1:]) """
            stufftowrite.append(sizeoutput)
            if pixelsize!=False:
                writesheet.write(0,col,'Area (um^2)')
            else:
                writesheet.write(0,col,'Area (pixels^2)')
            col+=1
        #print 'area'
        if 'DistanceTraveled' in i:
            speedoutput,dispoutput,integdistout,medspeed,maxdisp=e.speedanddisp(pixelsize,framerate)
            #print 'first'
            medspeedcell,meddispcell,maxspeedcell,maxdispcell,sumspeedx,sumspeedy,allmovementsums=e.speedanddisppercell(movielen,pixelsize,framerate)
            #print 'second'
            addonlist[2][identifier]=medspeed
            addonlist[3][identifier]=maxdisp
            addonlist[6][identifier]=[HXM.unziprezip(medspeedcell),HXM.unziprezip(meddispcell),HXM.unziprezip(maxspeedcell),HXM.unziprezip(maxdispcell),HXM.unziprezip(sumspeedx),HXM.unziprezip(sumspeedy)]
            #print addonlist[6][identifier]            
            """speedfilt=[]
            dispfilt=[]
            for j in speedoutput:
                speedfilt.append(j[1:])
            for k in dispoutput:
                dispfilt.append(k[1:])"""
            stufftowrite.append(speedoutput)
            stufftowrite.append(dispoutput)
            stufftowrite.append(integdistout)
            if pixelsize!=False:
                if framerate!=False:
                    writesheet.write(0,col,'Speed (um/second)')
                    writesheet.write(0,col+1,'Displacement from 0 (um)')
                    writesheet.write(0,col+2,'Integrated Distance (um)')
                else:
                    writesheet.write(0,col,'Speed (um/frame)')
                    writesheet.write(0,col+1,'Displacement from 0 (um)')
                    writesheet.write(0,col+2,'Integrated Distance (um)')
            else:
                if framerate!=False:
                    writesheet.write(0,col,'Speed (pixels/second)')
                    writesheet.write(0,col+1,'Displacement from 0 (pixels)')
                    writesheet.write(0,col+2,'Integrated Distance (pixels)')
                else:
                    writesheet.write(0,col,'Speed (pixels/frame)')
                    writesheet.write(0,col+1,'Displacement from 0 (pixels)')
                    writesheet.write(0,col+2,'Integrated Distance (pixels)')
            col+=4
    #print "intensity, area, distance"
    """outareaint=numpy.column_stack([medsize,medint])
    print 'areaint shape',numpy.shape(outareaint)
    numpy.savetxt(r'G:\\'+identifier+'Size.txt',outareaint)
    if pixelsize!=False:
        writesheet.write(0,col,'MSD (um^2)')
    else:
        writesheet.write(0,col,'MSD (pixels^2)')
    writesheet.write(0,col+1,'Disp SD')
    writesheet.write(0,col+2,'MSD n')
    msd={}
    for sp in dispoutput:
        for le in range(len(sp)):
            if type(sp[le])!=str:
                if le not in msd:
                    msd[le]=[sp[le]**2]
                else:
                    msd[le].append(sp[le]**2)
    #print msd
    addonlist[4][identifier]=[]
    for k in range(len(msd)):
        if k in msd:
            writesheet.write(1+k,col,numpy.mean(msd[k]))
            addonlist[4][identifier].append(numpy.mean(msd[k]))
            writesheet.write(1+k,col+1,numpy.std(msd[k]))
            writesheet.write(1+k,col+2,len(msd[k]))
    col+=4"""
    wheremsd=False
    if pixelsize!=False:
        if framerate!=False:
            wheremsd=col+1
            #col+=5
            msdeach,msdall=e.realmsd(pixelsize,framerate,identifier)
            writesheet.write(0,col,'Time(sec)')
            writesheet.write(0,col+1,'MSD')
            writesheet.write(0,col+2,'n of MSD')
            HXM.writesheet(writesheet,msdall,1,col)
            spotnum,diffco=e.diffco(pixelsize,framerate)
            addonlist[4][identifier]=[msdeach,HXM.unziprezip(msdall),diffco]
            writesheet.write(0,col+4,'Spot #')
            writesheet.write(0,col+5,'Diffusion Coefficient (um^2/sec)')
            HXM.writecol(writesheet,spotnum,col+4,1)
            HXM.writecol(writesheet,diffco,col+5,1)
            col+=6
        else:
            pass
    else:
        wheremsd=col+1
        #col+=5
        msdeach,msdall=e.realmsd(1,1,identifier)
        writesheet.write(0,col,'Time(frames)')
        writesheet.write(0,col+1,'MSD')
        writesheet.write(0,col+2,'n of MSD')
        HXM.writesheet(writesheet,msdall,1,col)
        addonlist[4][identifier]=[msdeach,HXM.unziprezip(msdall)]
        col+=4
    #print msdeach
    #print "MSD"
    addonlist[5][identifier]=[]
    if len(meanchannels)!=0:
        firstcolor=True
        writesheet.write(0,col,'Frame#')
        writesheet.write(0,col+1,'# particles')
        for entry in range(len(meanchannels)):
            writesheet.write(0,col+2*entry+2,'Mean per frame intensity-'+meanchannels[entry])
            writesheet.write(0,col+2*entry+3,'St.dev. per frame intensity-'+meanchannels[entry])
            frames,count,means,meandev=e.intmeanframe(meanchannels[entry])
            zipped=[meanchannels[entry]]+zip(frames,means)
            addonlist[5][identifier].append(zipped)
            if firstcolor==True:
                HXM.writecol(writesheet,frames,col,1)
                HXM.writecol(writesheet,count,col+1,1)
                firstcolor=False
            HXM.writecol(writesheet,means,col+2*entry+2,1)
            HXM.writecol(writesheet,meandev,col+2*entry+3,1)
        col+=2*len(meanchannels)+3
    if len(meanchannels)==2:
        try:
            frames,norm,normdev=e.intnormframe(meanchannels[0],meanchannels[1])
            zipped=[meanchannels[0]+'/'+meanchannels[1]]+zip(frames,norm)
            addonlist[5][identifier].append(zipped)
            writesheet.write(0,col,'Normalized mean intensity-'+meanchannels[0]+'/'+meanchannels[1])
            writesheet.write(0,col+1,'Normalized st.dev. intensity-'+meanchannels[0]+'/'+meanchannels[1])
            HXM.writecol(writesheet,norm,col,1)
            HXM.writecol(writesheet,normdev,col+1,1)
            col+=3
        except:
            frames,norm,normdev=e.intnormframe(meanchannels[1],meanchannels[0])
            zipped=[meanchannels[0]+'/'+meanchannels[1]]+zip(frames,norm)
            addonlist[5][identifier].append(zipped)
            writesheet.write(0,col,'Normalized mean intensity-'+meanchannels[1]+'/'+meanchannels[0])
            writesheet.write(0,col+1,'Normalized st.dev. intensity-'+meanchannels[1]+'/'+meanchannels[0])
            HXM.writecol(writesheet,norm,col,1)
            HXM.writecol(writesheet,normdev,col+1,1)
            col+=3
        
    else:
        col+=1
        
    #print "graph intensities"   
      
    hist=e.lengthhist()    
    writesheet.write(0,col+1,'Length Histogram')
    writesheet.write(1,col+1,'Length Bin')
    writesheet.write(1,col+2, '# of tracks')
    for ent in range(len(hist)):
        writesheet.write(2+ent,col+1,hist[ent][1])
        writesheet.write(2+ent,col+2,hist[ent][0])
        
    writesheet.write(0,col+4,'Frame #')
    writesheet.write(0,col+5,'# of spots')
    writesheet.write(0,col+6,'% Uncorrected X')
    writesheet.write(0,col+7,'% Uncorrected Y')
    HXM.writesheet(writesheet,allmovementsums,1,col+4)
    
    
        
    writeout=HXM.deepunziprezip(stufftowrite)
    #print writeout
    writesheetrows=2 #starting row value, increments each loop below
    extrasheets=1
    writesheet.write(0,0,'Spot Index')
    writesheet.write(0,1,'Frame #')
    writesheet.write(0,2,'X')
    writesheet.write(0,3,'Y')
#==============================================================================
#     if wheremsd!=False:
#         writesheet.write(0,wheremsd,'Time (sec)')
#         writesheet.write(0,wheremsd+1,'MSD')
#         writesheet.write(0,wheremsd+2,'n of MSD')
#==============================================================================
    for i in range(len(f)):
        if writesheetrows+len(f[i])>=65535:
            if len(identifier)>29:
                identifier=identifier[:29]
            writesheet=writebook.add_sheet(identifier+'_cont'+str(extrasheets))
            writesheetrows=1
            extrasheets+=1
        writesheet.write(writesheetrows-1,0,"Spot "+str(i+1)) #write the spot index
        writesheet.write(writesheetrows-1,1,f[i][0])
        HXM.writesheet(writesheet,f[i][1:],writesheetrows,1) #write the relative index, x and y info
        #print writeout[i][0]        
        HXM.writesheet(writesheet,writeout[i][1:],writesheetrows,4)
        """if wheremsd!=False:
            HXM.writesheet(writesheet,msdeach[i],writesheetrows,wheremsd)"""
        writesheetrows+=len(f[i])
    

    writesheet.flush_row_data()
    writebook.save(fileout)
    return addonlist

 
def dothestuff():
    """Asks the user for a .csv file, makes a copy of it into an excel file
    if that hasn't already been done, then creates an excel file that contains
    the x and y coordinates of all spots that pass the user's gate"""
   
    filefolder=eg.diropenbox(msg='Choose the parent folder that has all the subfolders containing tracking files')
    subfolds=[]
    for i in os.listdir(filefolder):
        if os.path.isdir(os.path.join(filefolder,i)):
            subfolds.append(i)
    foldstouse=eg.multchoicebox(msg='Which of these folders contain tracking files?',choices=subfolds)
    idents=[]
    for i in foldstouse:
        if 'out' in i:
            idents.append(i[:i.index('out')])
        elif 'Out' in i:
            idents.append(i[:i.index('Out')])
        else:
            idents.append(i)
    if len(idents)>1:
        baselineident=eg.choicebox(msg='Which of these is the baseline?', choices=idents)
    else:
        baselineident=None
    foldercsvs=[]
    for i in os.listdir(os.path.join(filefolder,foldstouse[0])):
        if '.csv' in i:
            foldercsvs.append(i)
    csvstouse=eg.multchoicebox(msg='Which of these files contain tracked objects?',choices=foldercsvs)
    addondicts=[{},{},{},{},{},{},{},{},{}]
    
    fulllen=int(eg.enterbox(msg='How many frames are in each original movie?'))   
    gate=eg.boolbox(msg='Do you want to discard tracks shorter than a given length?', choices=('Yes','No'))
    if gate:
        hasint=False
        setgate=eg.enterbox(msg='What is the minimum track length to consider?')
        while hasint==False: #Loop to make sure the user input is an integer
            try:
                setgate=int(setgate)
                hasint=True
            except:
                setgate=eg.enterbox(msg='What is the minimum track length to consider?')
    else:
        setgate=0
    maxgate=eg.ynbox(msg='Do you want to set a maximum track length to consider?')
    if maxgate==1:
        maxint=False
        while maxint==False:
            setmax=eg.enterbox(msg='What is the maximum track length to consider?')
            try:
                setmax=int(setmax)
                maxint=True
            except:
                maxint=False
    else:
        maxgate=fulllen
    pixelsize=eg.ynbox(msg='Do you want to convert from pixels to microns?')
    if pixelsize==1:
        pixfloat=False
        while pixfloat==False:
            pixelsize=eg.enterbox(msg='Enter the size of each pixel SIDE in microns')
            try:
                pixelsize=float(pixelsize)
                pixfloat=True
            except:
                pixfloat=False
    else:
        pixelsize=False
    framerate=eg.ynbox(msg='Do you want to convert from frames to seconds?')
    if framerate==1:
        framefloat=False
        while framefloat==False:
            framerate=eg.enterbox(msg='Enter the time between frames in seconds')
            try:
                framerate=float(framerate)
                framefloat=True
            except:
                framefloat=False
    else:
        framerate=False
        
    movies=eg.ynbox(msg='Do you want to make movies of the tracks?')
    suffixi=[]
    for i in os.listdir(os.path.join(filefolder,foldstouse[0])):
        if lastunderscoresuffix(i) not in suffixi:
            if 'PRJ' not in lastunderscoresuffix(i):
                suffixi.append(lastunderscoresuffix(i))
    drawsuffix=eg.choicebox(msg='Which of these is the suffix of the files you want to draw the tracks on?', choices=suffixi)

    
    reuse=eg.indexbox('Where to save the output?',choices=['Create a new file', 
                                                           'Add to a new sheet in another file',])
                                                           
    if reuse==0: #If the user wants to make a new file
        fileoutname=eg.filesavebox(msg='What do you want to name the file?',filetypes=["*.xls"])+'.xls'
        writebook=xlwt.Workbook() #make a new file
        writesheet=writebook.add_sheet('blank') #give it a sheet
        writebook.save(fileoutname)
    if reuse==1: #If the user wants a new sheet in an old file
        fileoutname=eg.fileopenbox(msg='Choose Excel file to write to',default='*.xls')

    for i in foldstouse:
        for j in csvstouse:
            #print i,j
            runtrackingcleanup(os.path.join(filefolder,i,j),setgate,maxgate,idents[foldstouse.index(i)]+j[(j.index('_')+1):-4],addondicts,fileoutname,movies,drawsuffix,fulllen,framerate,pixelsize)
            
    if framerate!=False:
        timeunits='seconds'
    else:
        timeunits='frames'
    #Make graphs
    if pixelsize!=False:
        distunits='(um)'
        if framerate!=False:
            speedunits='(um/sec)'
        else:
            speedunits='(um/frame)'
    else:
        distunits='(pixels)'
        if framerate!=False:
            speedunits='(pixels/sec)'
        else:
            speedunits='(pixels/frame)'
    reopenbook=xlrd.open_workbook(fileoutname)
    prevsheetnames=reopenbook.sheet_names()
    writegraphbook=copy(reopenbook)
    for oldsheets in range(len(prevsheetnames)-1):
        pulloldsheet=writegraphbook.get_sheet(oldsheets)
        pulloldsheet.flush_row_data()
        reopenbook.unload_sheet(oldsheets)
    graph=xlwt.easyxf('font: height 4000')
    mastergraphlist=[]
    masterbubblelist=[]
    speedhistdic={}
    disphistdic={}
    diffhistdic={}
    for i in foldstouse:
        graphstorun={}
        graphdict={'MSD':[],'IndScat':[],'AreaSpeedDist':[],'IntSpeedDist':[],'FrameVsInt':[],'Other':[]}
        #graphlist=[]
        bubblestorun={}
        bubblelist=[]
        for j in csvstouse:
            displist=None
            speedlist=None
            arealist=None
            chanlist=None
            comboident=idents[foldstouse.index(i)]+j[(j.index('_')+1):-4]
            whichcsv=j[(j.index('_')+1):-4]
            
            if comboident in addondicts[4].keys():
                graphstorun[(comboident,'Time ('+timeunits+')','MSD ('+distunits[1:-1]+'^2)')]=(addondicts[4][comboident][1][0],addondicts[4][comboident][1][1])
                #graphlist.append((whichcsv,'Frame #','MSD'))
                graphdict['MSD'].append((whichcsv,'Time ('+timeunits+')','MSD ('+distunits[1:-1]+'^2)'))
            if comboident in addondicts[6].keys():
                #print addondicts[6].keys()
                #print addondicts[6][comboident]
                graphstorun[(comboident,'Cell #', 'Median Speed'+speedunits)]=addondicts[6][comboident][0]
                graphdict['IndScat'].append((whichcsv,'Cell #', 'Median Speed'+speedunits))
                graphstorun[(comboident,'Cell #', 'Median Displacement'+distunits)]=addondicts[6][comboident][1]
                graphdict['IndScat'].append((whichcsv,'Cell #', 'Median Displacement'+distunits))
                graphstorun[(comboident,'Cell #', 'Maximum Speed'+speedunits)]=addondicts[6][comboident][2]
                graphdict['IndScat'].append((whichcsv,'Cell #', 'Maximum Speed'+speedunits))
                graphstorun[(comboident,'Cell #', 'Maximum Displacement'+distunits)]=addondicts[6][comboident][3]
                graphdict['IndScat'].append((whichcsv,'Cell #', 'Maximum Displacement'+distunits))
                graphstorun[(comboident,'Cell #', '% Uncorrected Movement in X')]=addondicts[6][comboident][4]
                graphdict['IndScat'].append((whichcsv,'Cell #', '% Uncorrected Movement in X'))
                graphstorun[(comboident,'Cell #', '% Uncorrected Movement in Y')]=addondicts[6][comboident][5]
                graphdict['IndScat'].append((whichcsv,'Cell #', '% Uncorrected Movement in Y'))
            if comboident in addondicts[5].keys():
                for color in addondicts[5][comboident]:
                    graphstorun[(comboident,'Frame #','Mean intensity('+color[0]+')')]=zip(*color[1:])
                    #graphlist.append((whichcsv,'Frame #','Mean intensity('+color[0]+')'))
                    graphdict['FrameVsInt'].append((whichcsv,'Frame #','Mean intensity('+color[0]+')'))
                #graphs about mean intensities at each frame                
                pass
            if comboident in addondicts[3].keys():
                displist=addondicts[3][comboident]
                if j[(j.index('_')+1):-4] not in disphistdic.keys():
                    disphistdic[j[(j.index('_')+1):-4]]=[[idents[foldstouse.index(i)]]+displist]
                else:
                    disphistdic[j[(j.index('_')+1):-4]].append([idents[foldstouse.index(i)]]+displist)
            if comboident in addondicts[2].keys():
                speedlist=addondicts[2][comboident]
                if j[(j.index('_')+1):-4] not in speedhistdic.keys():
                    speedhistdic[j[(j.index('_')+1):-4]]=[[idents[foldstouse.index(i)]]+speedlist]
                else:
                    speedhistdic[j[(j.index('_')+1):-4]].append([idents[foldstouse.index(i)]]+speedlist)
            if comboident in addondicts[4].keys():
                if framerate!=False:
                    if pixelsize!=False:
                        difflist=addondicts[4][comboident][2]
                        if j[(j.index('_')+1):-4] not in diffhistdic.keys():
                            diffhistdic[j[(j.index('_')+1):-4]]=[[idents[foldstouse.index(i)]]+difflist]
                        else:
                            diffhistdic[j[(j.index('_')+1):-4]].append([idents[foldstouse.index(i)]]+difflist)
            if comboident in addondicts[1].keys():
                arealist=addondicts[1][comboident]
            if arealist!=None:
                if speedlist!=None:
                    graphstorun[(comboident,'Median Area','Median Speed'+speedunits,'x')]=(arealist,speedlist)
                    #graphlist.append((whichcsv,'Median Area','Median Speed'))
                    graphdict['AreaSpeedDist'].append((whichcsv,'Median Area','Median Speed'+speedunits,'x'))
                if displist!=None:    
                    graphstorun[(comboident,'Median Area','Maximum Displacement'+distunits,'x')]=(arealist,displist)
                    #graphlist.append((whichcsv,'Median Area','Maximum Displacement'))
                    graphdict['AreaSpeedDist'].append((whichcsv,'Median Area','Maximum Displacement'+distunits,'x'))
            if comboident in addondicts[0].keys():
                chanidlist=[]
                chanlist=[]
                for k in range(len(addondicts[0][comboident])):
                    chanid=addondicts[0][comboident][k].pop(-1)
                    chanidlist.append(chanid)
                    chanlist.append(addondicts[0][comboident][k])
            if chanlist!=None:
                for k in range(len(chanidlist)):
                    if speedlist!=None:
                        graphstorun[(comboident,'Median Integrated Intensity-'+chanidlist[k],'Median Speed'+speedunits,'x')]=(chanlist[k],speedlist)
                        #graphlist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[k],'Median Speed'))
                        graphdict['IntSpeedDist'].append((whichcsv,'Median Integrated Intensity-'+chanidlist[k],'Median Speed'+speedunits,'x'))
                    if displist!=None:    
                        graphstorun[(comboident,'Median Integrated Intensity-'+chanidlist[k],'Maximum Displacement'+distunits,'x')]=(chanlist[k],displist)
                        #graphlist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[k],'Maximum Displacement'))
                        graphdict['IntSpeedDist'].append((whichcsv,'Median Integrated Intensity-'+chanidlist[k],'Maximum Displacement'+distunits,'x'))
                    if arealist!=None:
                        graphstorun[(comboident,'Median Area','Median Integrated Intensity-'+chanidlist[k],'xy')]=(arealist,chanlist[k])
                        #graphlist.append((whichcsv,'Median Area','Median Integrated Intensity-'+chanidlist[k]))
                        graphdict['Other'].append((whichcsv,'Median Area','Median Integrated Intensity-'+chanidlist[k],'xy'))
                if len(chanlist)>=2:
                    done=[]
                    for k in range(len(chanlist)):
                        for l in range(len(chanlist)):
                            if k!=l:
                                m=[k,l]
                                m.sort()
                                if m not in done:
                                    done.append(m)
                                    graphstorun[(comboident,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'xy')]=(chanlist[m[0]],chanlist[m[1]])
                                    #graphlist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]]))
                                    graphdict['Other'].append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'xy'))
                                    if speedlist!=None:
                                        bubblelist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Median Speed'+speedunits,'xy'))                                        
                                        bubblestorun[(comboident,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Median Speed'+speedunits,'xy')]=(chanlist[m[0]],chanlist[m[1]],speedlist)
                                    if displist!=None:
                                        bubblelist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Maximum Displacement'+distunits,'xy'))                                        
                                        bubblestorun[(comboident,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Maximum Displacement'+distunits,'xy')]=(chanlist[m[0]],chanlist[m[1]],displist)
                    
        graphlist=graphdict['MSD']+graphdict['IndScat']+graphdict['AreaSpeedDist']+graphdict['IntSpeedDist']+graphdict['FrameVsInt']+graphdict['Other']           
        #print graphlist        
        graphlistformaster=graphdict['MSD']+graphdict['AreaSpeedDist']+graphdict['IntSpeedDist']+graphdict['FrameVsInt']+graphdict['Other']
        writegraphsheet=writegraphbook.add_sheet(idents[foldstouse.index(i)]+'Graphs')
        for r in range(2):
            writegraphsheet.col(r).width=16500
        for r in range(((len(graphlist)+len(bubblelist)+4)/2)+1):
            writegraphsheet.row(r).set_style(graph)
        #graphlist.sort()    
        saveas=os.path.join(os.path.split(fileoutname)[0],'trash')
        graphcount=0
        outpdfs=PdfPages(fileoutname[:-4]+idents[foldstouse.index(i)]+'Graphs.pdf')
        for m in graphlist:
            for r in graphstorun:
                if m[0]==r[0][len(idents[foldstouse.index(i)]):]:
                    if m[1]==r[1]:
                        if m[2]==r[2]:
                            #print m, r
                            if 'MSD' in r[2]:
                                HXM.graphmsds([graphstorun[r][0]],[graphstorun[r][1]],saveas,mantitle=r[0]+'- MSD',xunits=timeunits,yunits=distunits,PDF=outpdfs)
                            else:
                                if len(r)>3:
                                    HXM.graphscatter([graphstorun[r][0]],[graphstorun[r][1]],[r[0]],saveas,r[1],r[2],mantitle=r[0]+' - '+r[1]+' vs. '+r[2],size=(450,330),PDF=outpdfs,log=r[3])
                                else:
                                    HXM.graphscatter([graphstorun[r][0]],[graphstorun[r][1]],[r[0]],saveas,r[1],r[2],mantitle=r[0]+' - '+r[1]+' vs. '+r[2],size=(450,330),PDF=outpdfs)
                            writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
                            os.remove(saveas+'.bmp')
                            os.remove(saveas+'.png')
                            graphcount+=1
        for m in bubblelist:
            for r in bubblestorun:
                if m[0]==r[0][len(idents[foldstouse.index(i)]):]:
                    if m[1]==r[1]:
                        if m[2]==r[2]:
                            if m[3]==r[3]:
                            #print m, r
                                if len(r)>4:
                                    HXM.graphbubble([bubblestorun[r][0]],[bubblestorun[r][1]],[bubblestorun[r][2]],[r[0]],saveas,r[1],r[2],r[3],mantitle=r[0]+' - '+r[1]+' vs. '+r[2]+' vs. '+r[3],size=(450,330),PDF=outpdfs,log=r[4])
                                else:
                                    HXM.graphbubble([bubblestorun[r][0]],[bubblestorun[r][1]],[bubblestorun[r][2]],[r[0]],saveas,r[1],r[2],r[3],mantitle=r[0]+' - '+r[1]+' vs. '+r[2]+' vs. '+r[3],size=(450,330),PDF=outpdfs)
                                writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
                                os.remove(saveas+'.bmp')
                                os.remove(saveas+'.png')
                                graphcount+=1
        HXM.graphmanyhists([speedlist],[comboident],comboident+'- Median Speed'+speedunits,saveas,size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1
        HXM.graphmanyhists([displist],[comboident],comboident+'- Maximum Displacement'+distunits,saveas,size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1
        if comboident in addondicts[4].keys():
            if framerate!=False:
                if pixelsize!=False:
                    HXM.graphmanyhists([difflist],[comboident],comboident+'- Diffusion Coefficient (um^2/sec)',saveas,size=(450,330),PDF=outpdfs)
                    writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                    graphcount+=1
        if comboident in addondicts[4].keys():
            #print addondicts[4][comboident][0]
            HXM.graphmsds(addondicts[4][comboident][0],addondicts[4][comboident][0],saveas,size=(450,330),PDF=outpdfs,xunits=timeunits,yunits=distunits,each=True)
            writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
            os.remove(saveas+'.bmp')
            os.remove(saveas+'.png')
            graphcount+=1
        if comboident in addondicts[7].keys():
            for m in addondicts[7][comboident].keys():
                HXM.graphtracksinacell(addondicts[7][comboident][m],saveas,'Integrated Distance Of Tracks'+distunits+'-'+comboident+' Cell '+str(m),PDF=outpdfs)
        
        mastergraphlist.append(graphstorun)
        masterbubblelist.append(bubblestorun)
        writegraphbook.save(fileoutname)
        outpdfs.close()
        writegraphsheet.flush_row_data()
    writegraphsheet=writegraphbook.add_sheet('Master Graphs')
    for r in range(2):
        writegraphsheet.col(r).width=16500
    for r in range(((len(mastergraphlist[0])+len(masterbubblelist[0])+3)/2)+1):
        writegraphsheet.row(r).set_style(graph)
    graphcount=0
    outpdfs=PdfPages(fileoutname[:-4]+'Graphs.pdf')
    for q in graphlistformaster:
        xlist=[]
        ylist=[]
        for z in range(len(foldstouse)):
            for m in mastergraphlist[z]:
                if q[0]==m[0][len(idents[z]):]:
                    #print 'q[0]', q[0], m[0]
                    if q[1]==m[1]:
                        #print 'q[1]', q[1], m[0]
                        if q[2]==m[2]:
                            #print 'q[2]', q[2], m[0]
                            xlist.append(mastergraphlist[z][m][0])
                            ylist.append(mastergraphlist[z][m][1])
        
        
        #print xlist, ylist
        if 'MSD' in q[2]:
            HXM.graphmsds(xlist,ylist,saveas,labels=idents,mantitle=q[0]+'- MSD',xunits=timeunits,yunits=distunits,PDF=outpdfs)
        else:
            if len(q)>3:                    
                HXM.graphscatter(xlist,ylist,idents,saveas,q[1],q[2],mantitle=q[0]+' - '+q[1]+' vs. '+q[2],size=(450,330),PDF=outpdfs,log=q[3])
            else:
                HXM.graphscatter(xlist,ylist,idents,saveas,q[1],q[2],mantitle=q[0]+' - '+q[1]+' vs. '+q[2],size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1
    
    for q in bubblelist:
        xlist=[]
        ylist=[]
        sizelist=[]
        for z in range(len(foldstouse)):
            for m in masterbubblelist[z]:
                if q[0]==m[0][len(idents[z]):]:
                    #print 'q[0]', q[0], m[0]
                    if q[1]==m[1]:
                        #print 'q[1]', q[1], m[0]
                        if q[2]==m[2]:
                            #print 'q[2]', q[2], m[0]
                            if q[3]==m[3]:
                                xlist.append(masterbubblelist[z][m][0])
                                ylist.append(masterbubblelist[z][m][1])
                                sizelist.append(masterbubblelist[z][m][2])
        
         
        #print xlist, ylist
        if len(q)>4:
            HXM.graphbubble(xlist,ylist,sizelist,idents,saveas,q[1],q[2],q[3],mantitle=q[0]+' - '+q[1]+' vs. '+q[2]+' vs. '+q[3],size=(450,330),PDF=outpdfs,log=q[4])
        else:
            HXM.graphbubble(xlist,ylist,sizelist,idents,saveas,q[1],q[2],q[3],mantitle=q[0]+' - '+q[1]+' vs. '+q[2]+' vs. '+q[3],size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1

       
    for r in speedhistdic.keys():
        speedhistdicidents=[] 
        speedhistvals=[]
        for q in speedhistdic[r]:
            #print q
            speedhistdicidents.append(q[0])
            speedhistvals.append(q[1:])
        if baselineident!=None:
            for k in range(len(speedhistdicidents)):
                if speedhistdicidents[k]==baselineident:
                    baseline=k
            #print speedhistdicidents
            for m in range(len(speedhistdicidents)):
                if m!=baseline:
                    ksvalue,pvalue=stats.ks_2samp(speedhistvals[baseline],speedhistvals[m])
                    speedhistdicidents[m]=speedhistdicidents[m]+'\n p='+'%.2f' %pvalue
        HXM.graphmanyhists(speedhistvals,speedhistdicidents,r+'- Median Speed'+speedunits,saveas,size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1 
        
       
    for r in disphistdic.keys():
        disphistdicidents=[] 
        disphistvals=[]
        for q in disphistdic[r]:
            disphistdicidents.append(q[0])
            disphistvals.append(q[1:])
        if baselineident!=None:
            for k in range(len(disphistdicidents)):
                if disphistdicidents[k]==baselineident:
                    baseline=k
            for m in range(len(disphistdicidents)):
                if m!=baseline:
                    ksvalue,pvalue=stats.ks_2samp(disphistvals[baseline],disphistvals[m])
                    disphistdicidents[m]=disphistdicidents[m]+'\n p='+'%.2f' %pvalue
        HXM.graphmanyhists(disphistvals,disphistdicidents,r+'- Maximum Displacement'+distunits,saveas,size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1
        
    for r in diffhistdic.keys():
        diffhistdicidents=[] 
        diffhistvals=[]
        for q in diffhistdic[r]:
            diffhistdicidents.append(q[0])
            diffhistvals.append(q[1:])
        if baselineident!=None:
            for k in range(len(diffhistdicidents)):
                if diffhistdicidents[k]==baselineident:
                    baseline=k
            for m in range(len(diffhistdicidents)):
                if m!=baseline:
                    ksvalue,pvalue=stats.ks_2samp(diffhistvals[baseline],diffhistvals[m])
                    diffhistdicidents[m]=diffhistdicidents[m]+'\n p='+'%.2f' %pvalue
        HXM.graphmanyhists(diffhistvals,diffhistdicidents,r+'- Diffusion Coefficient (um^2/sec)',saveas,size=(450,330),PDF=outpdfs)
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1
    writegraphsheet.flush_row_data()   
    outpdfs.close()
    writegraphbook.save(fileoutname)
    



if __name__=='__main__':
    again=True
    while again==True:
        dothestuff()
        a=eg.boolbox('Do you want to do another?',choices=['Yes','No'])
        if not a:
            again=False


