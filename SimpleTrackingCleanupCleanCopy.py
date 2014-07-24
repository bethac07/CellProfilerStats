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
from PIL import Image, ImageDraw
import os

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
    imageheadings=list(array3[0,:])
    for i in range(len(imageheadings)):
        if 'FileName' in imageheadings[i]:
            diffiles.append(imageheadings[i])        
    if len(diffiles)>1:
        filenameuse=eg.choicebox('Which of these is the image you want to draw on?',choices=diffiles)
        files=list(array3[:,imageheadings.index(filenameuse)])
        path=list(array3[:,imageheadings.index('Path'+filenameuse[4:])])
    else:
        for i in range(len(imageheadings)):
            if 'FileName' in imageheadings[i]:
                files=list(array3[:,i])
            if 'PathName' in imageheadings[i]:
                path=list(array3[:,i])
            
    fullfilename={}
    count=0
    for i in range(1,len(path)):
        if files[i]!=files[i-1]:
            count=0
        else:
            count+=1
        fullfilename[i]=(os.path.join(path[i],files[i]),count)
    return fullfilename
        
def runtrackingcleanup(filein,minlen,identifier,addonlist,fileout,makemovies,todrawsuffix,movielen):
    
    otherbook=xlrd.open_workbook(fileout) #Open it
    writebook=copy(otherbook) #make it writable
    writesheet=writebook.add_sheet(identifier) #add the new sheet

    e=Spots.spots(filein,movielen,minlen) #initialize the spot instances
    f=e.xyandt()
    spotheadings=e[0]
    names=imagefilenames(filein)
    
    if makemovies==True:
        #print 'movies'
        colorlist=['black','blue','crimson','darkorange','darkorchid','darkcyan','burlywood','deeppink','lawngreen','indigo','seashell','orangered','olive','powderblue','yellow']
    
        linedict={}
        circledict={}
        for i in range(len(f)):
            for j in range(1,len(f[i])-1):
                for z in range(int(j+f[i][0]),len(f[i])-1+int(f[i][0])):
                    if z in linedict.keys():
                        linedict[z].append(((f[i][j],f[i][j+1]),colorlist[i%15]))
                        circledict[z].append((f[i][j],colorlist[i%15]))
                    else:
                        linedict[z]=[((f[i][j],f[i][j+1]),colorlist[i%15])]
                        circledict[z]=[(f[i][j],colorlist[i%15])]
                        
            """for j in range(1,len(i)-1):
                draw.line((i[j],i[j+1]),fill='red')
            if i[0]==1:
                for j in range(1,len(i)-1):
                    draw.line((i[j],i[j+1]),fill='blue')
            draw.text((i[1][0]-5,i[1][1]-5),str(f.index(i)),fill='white',font=font)"""
        """for i in range(len(linedict)):
            if i in linedict.keys():
                print i, circledict[i],linedict[i]
            else:
                print i, 'none'"""
        
        for i in range(1,max(linedict.keys())):
            if i in linedict.keys():
                subfolder=os.path.split(names[i][0])[1]
                subfoldername=subfolder[:subfolder.index('.')]+os.path.split(filein)[1][os.path.split(filein)[1].index('_'):os.path.split(filein)[1].index('.')]+'_tracks'       
                newdir=os.path.join(os.path.split(filein)[0],subfoldername)
                if not os.path.isdir(newdir):
                    os.mkdir(newdir)
                #c=names[i][:names[i].index('.')]+'_'+str(i)+'_raw.bmp'
                c=os.path.join(os.path.split(filein)[0],os.path.split(names[i][0])[1][:os.path.split(names[i][0])[1].index('.')]+str(names[i][1])+todrawsuffix)
                #c=r'D:\ImageTesterOut\47A_03_R3D_D3D_PRJ-1_'+str(i)+'_TRF1.bmp'
                dd=Image.open(c)
                dd.save(os.path.join(newdir,subfolder[:subfolder.index('.')]+'_'+str(names[i][1])+'_rawmovies.jpg'))
                d=dd.convert('RGB')
                #import ImageDraw,ImageFont
                draw=ImageDraw.Draw(d)
            
                for j in circledict[i]:
                    #print j
                    makecircle=[(j[0][0]+3,j[0][1]+0),(j[0][0]+3,j[0][1]+1),(j[0][0]+2,j[0][1]+2),(j[0][0]+1,j[0][1]+3),(j[0][0]+0,j[0][1]+3),(j[0][0]-1,j[0][1]+3),(j[0][0]-2,j[0][1]+2),(j[0][0]-3,j[0][1]+1),(j[0][0]-3,j[0][1]+0),(j[0][0]-3,j[0][1]-1),(j[0][0]-2,j[0][1]-2),(j[0][0]-1,j[0][1]-3),(j[0][0]-0,j[0][1]-3),(j[0][0]+1,j[0][1]-3),(j[0][0]+2,j[0][1]-2),(j[0][0]+3,j[0][1]-1)]
                    draw.polygon(makecircle,outline=j[1])
                for k in linedict[i]:
                    draw.line(k[0],fill=k[1],width=2)
                del draw
                #print names[i][:names[i].index('.')]+'_'+str(i)+'movies.jpg'
                d.save(os.path.join(newdir,subfolder[:subfolder.index('.')]+'_'+str(names[i][1])+'movies.jpg'))
    
    col=4
    stufftowrite=[]
    meanchannels=[]
    for i in spotheadings:
        if 'MeanIntensity' in i:
            meanchannels.append(i[(i.index('MeanIntensity')+14):])
        if 'IntegratedIntensity' in i:
            channelname=i[(i.index('IntegratedIntensity')+20):]
            intoutput,medint=e.intint(channelname)
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
            sizeoutput,medsize=e.size()
            addonlist[1][identifier]=medsize
            """sizefilt=[]
            for j in sizeoutput:
                sizefilt.append(j[1:]) """
            stufftowrite.append(sizeoutput)
            writesheet.write(0,col,'Area')
            col+=1
        if 'DistanceTraveled' in i:
            speedoutput,dispoutput,medspeed,maxdisp=e.speedanddisp()
            addonlist[2][identifier]=medspeed
            addonlist[3][identifier]=maxdisp
            """speedfilt=[]
            dispfilt=[]
            for j in speedoutput:
                speedfilt.append(j[1:])
            for k in dispoutput:
                dispfilt.append(k[1:])"""
            stufftowrite.append(speedoutput)
            stufftowrite.append(dispoutput)
            writesheet.write(0,col,'Speed')
            writesheet.write(0,col+1,'Displacement from 0')
            col+=3

    
    writesheet.write(0,col,'MSD')
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
    addonlist[4][identifier]=[]
    for k in range(len(msd)):
        if k in msd:
            writesheet.write(1+k,col,numpy.mean(msd[k]))
            addonlist[4][identifier].append(numpy.mean(msd[k]))
            writesheet.write(1+k,col+1,numpy.std(msd[k]))
            writesheet.write(1+k,col+2,len(msd[k]))
    col+=4
    
    addonlist[5][identifier]=[]
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
        
        
    

    hist=e.lengthhist()    
    writesheet.write(0,col+1,'Length Histogram')
    writesheet.write(1,col+1,'Length Bin')
    writesheet.write(1,col+2, '# of tracks')
    for ent in range(len(hist)):
        writesheet.write(2+ent,col+1,hist[ent][1])
        writesheet.write(2+ent,col+2,hist[ent][0])
        
    writeout=HXM.deepunziprezip(stufftowrite)
    writesheetrows=2 #starting row value, increments each loop below
    for i in range(len(f)):
        writesheet.write(writesheetrows-1,0,"Spot "+str(i+1)) #write the spot index
        writesheet.write(writesheetrows-1,1,f[i][0])
        HXM.writesheet(writesheet,f[i][1:],writesheetrows,2) #write the x and y info
        HXM.writecol(writesheet,range(len(f[i])-1),1,writesheetrows) #write the time index
        #print writeout[i][0]        
        HXM.writesheet(writesheet,writeout[i][1:],writesheetrows,4)        
        writesheetrows+=len(f[i])

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
    addondicts=[{},{},{},{},{},{}]
    
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
        
    movies=eg.ynbox(msg='Do you want to make movies of the tracks?')
    if movies:
        suffixi=[]
        for i in os.listdir(os.path.join(filefolder,foldstouse[0])):
            if lastunderscoresuffix(i) not in suffixi:
                if 'PRJ' not in lastunderscoresuffix(i):
                    suffixi.append(lastunderscoresuffix(i))
        drawsuffix=eg.choicebox(msg='Which of these is the suffix of the files you want to draw the tracks on?', choices=suffixi)
    else:
        drawsuffix=None
    
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
            runtrackingcleanup(os.path.join(filefolder,i,j),setgate,idents[foldstouse.index(i)]+j[(j.index('_')+1):-4],addondicts,fileoutname,movies,drawsuffix,fulllen)
            
          
    #Make graphs
    reopenbook=xlrd.open_workbook(fileoutname)
    writegraphbook=copy(reopenbook)
    graph=xlwt.easyxf('font: height 4000')
    mastergraphlist=[]
    masterbubblelist=[]
    speedhistdic={}
    disphistdic={}
    for i in foldstouse:
        graphstorun={}
        graphlist=[]
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
                graphstorun[(comboident,'Frame #','MSD')]=(range(len(addondicts[4][comboident])),addondicts[4][comboident])
                graphlist.append((whichcsv,'Frame #','MSD'))
            if comboident in addondicts[5].keys():
                for color in addondicts[5][comboident]:
                    graphstorun[(comboident,'Frame #','Mean intensity('+color[0]+')')]=zip(*color[1:])
                    graphlist.append((whichcsv,'Frame #','Mean intensity('+color[0]+')'))
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
            if comboident in addondicts[1].keys():
                arealist=addondicts[1][comboident]
            if arealist!=None:
                if speedlist!=None:
                    graphstorun[(comboident,'Median Area','Median Speed')]=(arealist,speedlist)
                    graphlist.append((whichcsv,'Median Area','Median Speed'))
                if displist!=None:    
                    graphstorun[(comboident,'Median Area','Maximum Displacement')]=(arealist,displist)
                    graphlist.append((whichcsv,'Median Area','Maximum Displacement'))
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
                        graphstorun[(comboident,'Median Integrated Intensity-'+chanidlist[k],'Median Speed')]=(chanlist[k],speedlist)
                        graphlist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[k],'Median Speed'))
                    if displist!=None:    
                        graphstorun[(comboident,'Median Integrated Intensity-'+chanidlist[k],'Maximum Displacement')]=(chanlist[k],displist)
                        graphlist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[k],'Maximum Displacement'))
                    if arealist!=None:
                        graphstorun[(comboident,'Median Area','Median Integrated Intensity-'+chanidlist[k])]=(arealist,chanlist[k])
                        graphlist.append((whichcsv,'Median Area','Median Integrated Intensity-'+chanidlist[k]))
                if len(chanlist)>=2:
                    done=[]
                    for k in range(len(chanlist)):
                        for l in range(len(chanlist)):
                            if k!=l:
                                m=[k,l]
                                m.sort()
                                if m not in done:
                                    done.append(m)
                                    graphstorun[(comboident,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]])]=(chanlist[m[0]],chanlist[m[1]])
                                    graphlist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]]))
                                    if speedlist!=None:
                                        bubblelist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Median Speed'))                                        
                                        bubblestorun[(comboident,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Median Speed')]=(chanlist[m[0]],chanlist[m[1]],speedlist)
                                    if displist!=None:
                                        bubblelist.append((whichcsv,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Maximum Displacement'))                                        
                                        bubblestorun[(comboident,'Median Integrated Intensity-'+chanidlist[m[0]],'Median Integrated Intensity-'+chanidlist[m[1]],'Maximum Displacement')]=(chanlist[m[0]],chanlist[m[1]],displist)
                    
                    
        writegraphsheet=writegraphbook.add_sheet(idents[foldstouse.index(i)]+'Graphs')
        for r in range(2):
            writegraphsheet.col(r).width=16500
        for r in range(((len(graphlist)+len(bubblelist))/2)+1):
            writegraphsheet.row(r).set_style(graph)
        graphlist.sort()    
        saveas=os.path.join(os.curdir,'trash')
        graphcount=0
        for m in graphlist:
            for r in graphstorun:
                if m[0]==r[0][len(idents[foldstouse.index(i)]):]:
                    if m[1]==r[1]:
                        if m[2]==r[2]:
                            #print m, r
                            HXM.graphscatter([graphstorun[r][0]],[graphstorun[r][1]],[r[0]],saveas,r[1],r[2],mantitle=r[0]+' - '+r[1]+' vs. '+r[2],size=(450,330))
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
                                HXM.graphbubble([bubblestorun[r][0]],[bubblestorun[r][1]],[bubblestorun[r][2]],[r[0]],saveas,r[1],r[2],r[3],mantitle=r[0]+' - '+r[1]+' vs. '+r[2]+' vs. '+r[3],size=(450,330))
                                writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
                                os.remove(saveas+'.bmp')
                                os.remove(saveas+'.png')
                                graphcount+=1
        
        mastergraphlist.append(graphstorun)
        masterbubblelist.append(bubblestorun)
        writegraphbook.save(fileoutname)
    
    writegraphsheet=writegraphbook.add_sheet('Master Graphs')
    for r in range(2):
        writegraphsheet.col(r).width=16500
    for r in range(((len(mastergraphlist[0])+len(masterbubblelist[0])+(2*len(speedhistdic.keys())))/2)+1):
        writegraphsheet.row(r).set_style(graph)
    graphcount=0
    
    for q in graphlist:
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
        HXM.graphscatter(xlist,ylist,idents,saveas,q[1],q[2],mantitle=q[0]+' - '+q[1]+' vs. '+q[2],size=(450,330))
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
        HXM.graphbubble(xlist,ylist,sizelist,idents,saveas,q[1],q[2],q[3],mantitle=q[0]+' - '+q[1]+' vs. '+q[2]+' vs. '+q[3],size=(450,330))
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
                    speedhistdicidents[m]=speedhistdicidents[m]+' p='+'%.2f' %pvalue
        HXM.graphmanyhists(speedhistvals,speedhistdicidents,r+'- Median Speed',saveas,size=(450,330))
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
                    disphistdicidents[m]=disphistdicidents[m]+' p='+'%.2f' %pvalue
        HXM.graphmanyhists(disphistvals,disphistdicidents,r+'- Maximum Displacement',saveas,size=(450,330))
        writegraphsheet.insert_bitmap(saveas+'.bmp',graphcount/2,graphcount%2)
        os.remove(saveas+'.bmp')
        os.remove(saveas+'.png')
        graphcount+=1
    
       
    
    writegraphbook.save(fileoutname)



if __name__=='__main__':
    again=True
    while again==True:
        dothestuff()
        a=eg.boolbox('Do you want to do another?',choices=['Yes','No'])
        if not a:
            again=False


