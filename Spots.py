# -*- coding: utf-8 -*-
"""
Created on Fri Mar 04 14:31:25 2011

@author: Beth Cimini
"""
import numpy
import os
import random
from scipy import stats
from HandyXLModules import pythag



class spots(list):
    def __init__(self,filename,movielen,minlength=0,maxlen=100000):
        """Read an Excel file, pick out all the 
        instances of a given spot (as identified by Label)
        and give them a list of all the attributes measured
        The overall container is a list, with the headings as
        index 0 and the rest of the spots in order as lists of lists
        (one list for each timepoint, containing all the data
        from that timepoint)
        minlength allows the user to gate out instances that 
        don't persist for a certain length of time"""
        
        
        array=numpy.genfromtxt(filename,delimiter=',',dtype=None)
        array2=numpy.genfromtxt(filename,delimiter=',',dtype=numpy.float64)
        headings=list(array[0,:]) #read headings
        
        
        self.append(headings) #save headings to self
        for i in range(len(headings)): #identify the label column
            if 'TrackObjects_Label' in headings[i]:
                labelcolumn=i
            if 'ImageNumber' in headings[i]:
                imnumcolumn=i
            if 'TrackObjects_Lifetime' in headings[i]:
                lifecolumn=i
        
        alllabels=list(array2[:,labelcolumn]) #a list of all the spot instances at all timepoints
        imnumlist=list(array2[:,imnumcolumn])
        lifelist=list(array2[:,lifecolumn])
        for w in range(1,len(alllabels)):
            alllabels[w]=int(alllabels[w])
            imnumlist[w]=int(imnumlist[w])
            lifelist[w]=int(lifelist[w])
        starts=[1]
        stops=[]        
        for z in range(2,max(imnumlist[1:])):
            c=[]
            for i in range(1,len(imnumlist)):
                if imnumlist[i]==z:
                    c.append([i,alllabels[i],lifelist[i]])
            if len(c)>0:
                first=True
                while first==True:
                    if c[0][1]!=1:
                        first=False
                    for j in range(len(c)):
                        if c[j][2]!=0:
                            first=False
                        """if c[j][0]-c[j-1][0]!=1:
                            first=False"""
                    if first==True:
                        starts.append(c[0][0])
                        stops.append(c[0][0])
                        first =False
        stops.append(len(imnumlist))
        #print starts, stops
        
        for m in range(len(starts)):
            sublabels=alllabels[starts[m]:stops[m]]
            sublife=lifelist[starts[m]:stops[m]]
            """Add screening for only one telomere per identifier"""
            subarray=array2[starts[m]:stops[m],:]
            try:   
                maxlabel=max(sublabels) #find the number of spot instances
                for r in range(1,(maxlabel+1)): #for each spot identifier
                    if sublabels.count(r)>=minlength: #checks if the spot lasts long enough
                        a=[]  #the list of all information for an individual spot
                        setcount=0
                        for k in range(len(sublabels)): #if so, add all the information about that spot
                                                        #at each timepoint
                            if sublabels[k]==r:
                                if sublife[k]==setcount:
                                    #while setcount<=maxlen:
                                    a.append([(int(subarray[k][0])-1)%movielen]+list(subarray[k])[1:]+[int(subarray[k][0])])
                                    setcount+=1
                        self.append(a) #add the compiled spot information to the master list
                        #print starts[m],r,'added'
            except:
                continue
    
    
    def realmsd(self,pixelsize,framerate,identifier):
        """because apparently I am dumb"""
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        
        tracklist=[]
        alltracks={}
        count=0
        trackcount=0
        maxlencount=[]
        for i  in self[1:]:
            trackcount+=1
            maxlencount.append(len(i))
        for i  in self[1:]:
            trackframepos={}
            startframe=i[0][-1]
            for frame in range(len(i)):
                trackframepos[i[frame][-1]-startframe]=(i[frame][x],i[frame][y])
            allframes=trackframepos.keys()
            #print 'allframes',allframes
            allframes.sort()
            squareddisp={}
            for sep in range(1,allframes[-1]-5):
                squareddisp[sep]=[]
                for j in allframes:
                    if j+sep in allframes:
                        squareddisp[sep].append((pixelsize*pythag(trackframepos[j+sep][0],trackframepos[j][0],trackframepos[j+sep][1],trackframepos[j][1]))**2)
            #print squareddisp
            allsep=squareddisp.keys()
            allsep.sort()
            for i in allsep:
                if len(squareddisp[i])==0:
                    allsep=allsep[:allsep.index(i)]
                    break
            trackfinal=[]
            for eachsep in allsep:
                trackfinal.append((eachsep*framerate,numpy.mean(squareddisp[eachsep]),len(squareddisp[eachsep])))
                if eachsep not in alltracks.keys():
                    alltracks[eachsep]=squareddisp[eachsep]
                else:
                    alltracks[eachsep]+=squareddisp[eachsep]
                #print trackfinal
            if len(trackfinal)>0:
                tracklist.append(trackfinal)
            count+=1
        finalsep=alltracks.keys()
        finalsep.sort()
        alltracklist=[]
        for eachframe in finalsep:
            alltracklist.append((eachframe*framerate,numpy.mean(alltracks[eachframe]),len(alltracks[eachframe])))
        return tracklist, alltracklist
        


    def xyandrealt(self):
        """Pull the x and y coordinates for each spot at each 
        timepoint- index 0 gives the user the frame at which
        the spot first appeared, and the rest are tuples
        containing the coordinate information for each
        timepoint"""
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        
        a=[]
        for i in self[1:]:
            b=[i[0][-1]]
            for j in i:
                b.append((j[-1]-b[0],j[x],j[y]))
            a.append(b)
        #print a
        return a

    def trackstonow(self,movielen):
        """Description"""
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        
        count=0
        tracksdict={}
        tracksonlyfinaldict={}
        for i in self[1:]:
            trackinprog=[[],[],count%7]
            cell=((i[0][-1]-1)/movielen)+1
            zlist=[]
            relzlist=[]
            lastframe=i[-1][-1]
            for j in i:
                zlist.append(j[-1])
                relzlist.append(j[-1]%200)
                trackinprog[0]=list(trackinprog[0])+[j[x]]
                trackinprog[1]=list(trackinprog[1])+[j[y]]
                if j[-1] not in tracksdict.keys():
                    tracksdict[j[-1]]=[list(trackinprog)]
                else:
                    tracksdict[j[-1]].append(list(trackinprog))
                if j[-1]==lastframe:
                    zlist=list(relzlist)
                    if cell not in tracksonlyfinaldict.keys():
                        tracksonlyfinaldict[cell]=[list(trackinprog)[:2]+[relzlist]]
                        #tracksonlyfinaldict[cell]=[[map(lambda x:x+1,trackinprog[0]),map(lambda x:x+1,trackinprog[1]),zlist]]
                    else:
                        tracksonlyfinaldict[cell].append(list(trackinprog)[:2]+[relzlist])
                        #tracksonlyfinaldict[cell].append([map(lambda x:x+1,trackinprog[0]),map(lambda x:x+1,trackinprog[1]),zlist])
                #print j[-1], trackinprog,tracksdict
            count+=1
        #print tracksdict[1],tracksdict[2],tracksdict[3]
        return tracksdict,tracksonlyfinaldict
 
    def intint(self,channelname):
        for i in self[0]:
            if 'Intensity_IntegratedIntensity' in i:
                if channelname in i:
                   k=self[0].index(i)
        a=[]
        medvals=[]
        for i in self[1:]:
            b=[i[0][0]]
            for j in i:
                b.append(j[k])
            a.append(b)
            medvals.append(numpy.median(b))
        return a, medvals

    def intmeanframe(self,channelname):
        for i in self[0]:
            if 'Intensity_MeanIntensity' in i:
                if channelname in i:
                   k=self[0].index(i)
        frames={}
        for i in self[1:]:
            for j in i:
                if j[0] not in frames.keys():
                    frames[j[0]]=[j[k]]
                else:
                    frames[j[0]]+=[j[k]]
        b=frames.keys()
        b.sort()
        meanframes=[]
        stdevframes=[]
        n=[]
        for i in b:
            n.append(len(frames[i]))
            meanframes.append(numpy.mean(frames[i]))
            stdevframes.append(numpy.std(frames[i]))
        return b,n,meanframes,stdevframes
        
    def intnormframe(self,channel1name,channel2name):
        for i in self[0]:
            if 'Intensity_MeanIntensity' in i:
                if channel1name in i:
                    k=self[0].index(i)
                if channel2name in i:
                    m=self[0].index(i)
        frames={}
        for i in self[1:]:
            for j in i:
                if j[0] not in frames.keys():
                    frames[j[0]]=[j[k]/j[m]]
                else:
                    frames[j[0]]+=[j[k]/j[m]]
        b=frames.keys()
        b.sort()
        meanframes=[]
        stdevframes=[]
        for i in b:
            meanframes.append(numpy.mean(frames[i]))
            stdevframes.append(numpy.std(frames[i]))
        return b,meanframes,stdevframes
    
    def size(self,pixelsize):
        for heading in self[0]:
            if 'AreaShape_Area' in heading:
                area=self[0].index(heading)
        
        if pixelsize==False:
            pixelsize=1
        a=[]
        medvals=[]
        for i in self[1:]:
            b=[i[0][0]]
            for j in i:
                b.append(j[area]*(pixelsize**2))
            a.append(b)
            medvals.append(numpy.median(b))
        return a,medvals

    def diffco(self,pixelsize,framerate):
        for heading in self[0]:
            if 'DistanceTraveled' in heading:
                dist=self[0].index(heading)
        spotnum=[]
        calcdiffco=[]
        count=1
        for i in self[1:]:
            startframe=i[0][-1]
            lastframe=i[-1][-1]
            distinpix=i[-1][dist]*pixelsize
            diff=(distinpix**2)/(4*framerate*(lastframe+1-startframe))
            spotnum.append(count)
            calcdiffco.append(diff)
            count+=1
        return spotnum,calcdiffco
            

    def speedanddisp(self,pixelsize,framerate):
        for heading in self[0]:
            if 'DistanceTraveled' in heading:
                dist=self[0].index(heading)
            if 'TrajectoryX' in heading:
                trajx=self[0].index(heading)
            if 'TrajectoryY' in heading:
                trajy=self[0].index(heading)
        if pixelsize==False:
            pixelsize=1
        if framerate==False:
            framerate=1
        speed=[]
        disp=[]
        integ=[]
        medspeed=[]
        maxdisp=[]
        for i in self[1:]:
            currframe=i[0][-1]
            instdisp=[0]
            instspeed=[0]
            instintegdist=[0]
            instspeednumonly=[0]
            for j in range(1,len(i)):
                pythag=((i[j][trajx]**2)+(i[j][trajy]**2))**(0.5)
                instdisp.append(i[j][dist]*pixelsize)
                instintegdist.append(instintegdist[-1]+pythag*pixelsize)
                if i[j][-1]-currframe==1:
                    instspeed.append(pythag*pixelsize/framerate)
                    instspeednumonly.append(pythag*pixelsize/framerate)
                else:
                    instspeed.append('')
                currframe=i[j][-1]
            disp.append(instdisp)
            speed.append(instspeed)
            integ.append(instintegdist)
            maxdisp.append(numpy.max(instdisp))
            medspeed.append(numpy.median(instspeednumonly))
            
            
        return speed,disp,integ,medspeed,maxdisp
        
    def speedanddisppercell(self,movielen,pixelsize,framerate):
        if pixelsize==False:
            pixelsize=1
        if framerate==False:
            framerate=1
        for heading in self[0]:
            if 'DistanceTraveled' in heading:
                dist=self[0].index(heading)
            if 'TrajectoryX' in heading:
                trajx=self[0].index(heading)
            if 'TrajectoryY' in heading:
                trajy=self[0].index(heading)
        
        medspeed=[]
        meddisp=[]
        maxspeed=[]
        maxdisp=[]
        allsumspeedx=[]
        allsumspeedy=[]
        sumspeed={}
        allpercentsums=[]
        for i in self[1:]:
            cell=((i[0][-1]-1)/movielen)+1
            currframe=i[0][-1]
            disp=[0]
            speed=[0]
            scat=cell+random.uniform(-.25,.25)
            for j in range(1,len(i)):
                disp.append(i[j][dist]*pixelsize)
                if i[j][-1]-currframe==1:
                    speed.append((((i[j][trajx]**2)+(i[j][trajy]**2))**(0.5))*pixelsize/framerate)
                if (cell,currframe) not in sumspeed.keys():
                    sumspeed[(cell,currframe)]=[i[j][trajx]*pixelsize/framerate,abs(i[j][trajx])*pixelsize/framerate,i[j][trajy]*pixelsize/framerate,abs(i[j][trajy])*pixelsize/framerate,1]
                else:
                    #print currframe, sumspeed[(cell,currframe)], i[j][trajx], i[j][trajy]
                    sumspeed[(cell,currframe)][0]+=i[j][trajx]*pixelsize/framerate
                    sumspeed[(cell,currframe)][1]+=abs(i[j][trajx])*pixelsize/framerate
                    sumspeed[(cell,currframe)][2]+=i[j][trajy]*pixelsize/framerate
                    sumspeed[(cell,currframe)][3]+=abs(i[j][trajy])*pixelsize/framerate
                    sumspeed[(cell,currframe)][4]+=1
                    #print sumspeed[(cell,currframe)]
                currframe=i[j][-1]
            medspeed.append((scat,numpy.median(speed)))
            meddisp.append((scat,numpy.median(disp)))
            maxspeed.append((scat,numpy.max(speed)))
            maxdisp.append((scat,numpy.max(disp)))
        #print sumspeed
        for i in sumspeed.keys():
            scat=i[0]+random.uniform(-.25,.25)
            allsumspeedx.append((scat,sumspeed[i][0]*100/sumspeed[i][1]))
            allsumspeedy.append((scat,sumspeed[i][2]*100/sumspeed[i][3]))
            allpercentsums.append((i[1],sumspeed[i][4],sumspeed[i][0]*100/sumspeed[i][1],sumspeed[i][2]*100/sumspeed[i][3]))
        allpercentsums.sort()
        return medspeed,meddisp,maxspeed,maxdisp,allsumspeedx,allsumspeedy,allpercentsums
        
    def intintandintdistpercell(self,movielen,channelname,pixelsize):
        for heading in self[0]:
            if 'Intensity_IntegratedIntensity' in heading:
                if channelname in heading:
                   intint=self[0].index(heading)
            if 'TrajectoryX' in heading:
                trajx=self[0].index(heading)
            if 'TrajectoryY' in heading:
                trajy=self[0].index(heading)
        if pixelsize==False:
            pixelsize=1
        tracksbycell={}    
        intensitybycell={}
        classifiedbycell={}
        
        for i in self[1:]:
            cell=((i[0][-1]-1)/movielen)+1
            if cell not in tracksbycell.keys():
                tracksbycell[cell]=[]
                intensitybycell[cell]=[]
            if i[0][-1]%movielen==1:    
                disp=[0]
                intensity=[]
                frame=[0]    
                for j in range(1,len(i)):
                    frame.append(i[j][-1]-i[0][-1])
                    intensity.append(i[j][intint])
                    pythag=((i[j][trajx]**2)+(i[j][trajy]**2))**(0.5)
                    disp.append(disp[-1]+pythag*pixelsize)
                intensitybycell[cell].append(numpy.median(intensity))
                tracksbycell[cell].append([numpy.median(intensity),frame,disp])
                
        for eachcell in intensitybycell.keys():
            if len(intensitybycell[eachcell])>3:
                classifiedbycell[eachcell]=[]
                per25=stats.scoreatpercentile(intensitybycell[eachcell],25)
                per50=stats.scoreatpercentile(intensitybycell[eachcell],50)
                per75=stats.scoreatpercentile(intensitybycell[eachcell],75)
                for eachtrack in tracksbycell[eachcell]:
                    if eachtrack[0]<per25:
                        classifiedbycell[eachcell].append([0]+eachtrack[1:])
                    elif eachtrack[0]<per50:
                        classifiedbycell[eachcell].append([1]+eachtrack[1:])
                    elif eachtrack[0]<per75:
                        classifiedbycell[eachcell].append([2]+eachtrack[1:])
                    else:
                        classifiedbycell[eachcell].append([3]+eachtrack[1:])
            else:
                classifiedbycell[eachcell]=[]
                for eachtrack in tracksbycell[eachcell]:
                    classifiedbycell[eachcell].append([3]+eachtrack[1:])

        return classifiedbycell
        
    def xyandtindiv(self,indiv):
        """Pull the x and y coordinates of a single spot.
        Index 0 is the frame number of the first appearance"""
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        a=[self[indiv][0][0]]
        for j in self[indiv]:
            a.append((j[x],j[y]))
        return a
    
    def lengthhist(self):
        """Provide a rough length histogram for the spots"""
        pulllengths=[]
        for i in self[1:]:
            pulllengths.append(len(i)-1)
        lengthhist,bins=numpy.histogram(pulllengths,bins=20)
        lengthout=[]
        for j in range(len(lengthhist)):
            binstr=str(bins[j])+'-'+str(bins[j+1])
            lengthout.append((lengthhist[j],binstr))
        return lengthout
        
if __name__=='__main__':
    e=spots(r'G:\TrackingOutputFolders200Frames\RPE-CRISPR\DO_Telomeres.csv',200,minlength=50)
    f=e.realmsd(.156,.2,'RPECRISPR')