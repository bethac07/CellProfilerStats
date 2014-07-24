# -*- coding: utf-8 -*-
"""
Created on Fri Mar 04 14:31:25 2011

@author: Beth Cimini
"""
import numpy
import os



class spots(list):
    def __init__(self,filename,movielen,minlength=0):
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
            """Add screening for only one telomere per image"""
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
                                    a.append([(int(subarray[k][0])-1)%movielen]+list(subarray[k])[1:])
                                    setcount+=1
                        self.append(a) #add the compiled spot information to the master list
                        #print starts[m],r,'added'
            except:
                continue
    
    

    def xyandt(self):
        """Pull the x and y coordinates for each spot at each 
        timepoint- index 0 gives the user the frame at which
        the spot first appeared, and the rest are tuples
        containing the coordinate information for each
        timepoint"""
        x=self[0].index('Location_Center_X')
        y=self[0].index('Location_Center_Y')
        a=[]
        for i in self[1:]:
            b=[i[0][0]]
            for j in i:
                b.append((j[x],j[y]))
            a.append(b)
        return a

 
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
    
    def size(self):
        for heading in self[0]:
            if 'AreaShape_Area' in heading:
                area=self[0].index(heading)
        a=[]
        medvals=[]
        for i in self[1:]:
            b=[i[0][0]]
            for j in i:
                b.append(j[area])
            a.append(b)
            medvals.append(numpy.median(b))
        return a,medvals

    def speedanddisp(self):
        for heading in self[0]:
            if 'DistanceTraveled' in heading:
                dist=self[0].index(heading)
            if 'IntegratedDistance' in heading:
                intdist=self[0].index(heading)
            if 'Linearity' in heading:
                lin=self[0].index(heading)
        
        a=[]
        z=[]
        medspeed=[]
        maxdisp=[]
        for i in self[1:]:
            b=[0]
            y=[0]
            for j in i:
                b.append(j[dist])
                if j[intdist]*j[lin]!=j[intdist]*j[lin]:
                    y.append('x')
                else:
                    y.append(float(j[intdist])*float(j[lin]))
            a.append(b)
            medspeed.append(numpy.median(b))
            z.append(y)
            y1=[]
            for j in y:
                try:
                    j=float(j)
                    y1.append(j)
                except:
                    pass
            maxdisp.append(numpy.max(y1))
        return a,z,medspeed,maxdisp
        
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
        
