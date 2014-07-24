"""CSV Combiner- part of CellProfiler Stats- by Beth Cimini Sept 2010

    Takes the directory of .csv files created by the measurement algorithms of CellProfiler and combines them into
    a single Excel file, adding a nearest neighbor analysis for up to four types of objects if requested."""

import xlrd
import xlwt
from xlutils.copy import copy
import os
import csv
import math
import easygui as eg



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
    for y in filelist: #for each csv file...
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
    return k

def pull2vals(workbook,sheetindex,firstvalue="Location_Center_X",secondvalue="Location_Center_Y"): #firstvalue and secondvalue can be changed if a different default is desired
    #Pull two columns of values, either an x and y for nearest neighbor analysis or two values to
    #be directly compared (ie size vs intensity, nearest neighbor vs intensity)
    wb=xlrd.open_workbook(workbook,on_demand=True) #Open a particular excelfile
    c=[]
    para1=wb.sheet_by_index(sheetindex) #open the particular subsheet
    p1cols=[]
    for m in range(para1.ncols): #look for the columns that contain the parameters of interest
        if para1.cell(0,m).value==firstvalue:
            p1cols.append(m)
        elif para1.cell(0,m).value==secondvalue:
            p1cols.append(m)
    
    b= int(para1.cell(para1.nrows-1,0).value) #Determine how many images- the number in the first column of the last row

    #For each image, pull all the values of its objects into a single list
    for i in range(1,(b+1)): #Deal with each image separately
        d=[i]
        for j in range(1,para1.nrows):  #Over all objects...
            if para1.cell(j,0).value==i: #If the object is part of the image we're looking for...
                d.append([para1.cell(j,p1cols[0]).value,para1.cell(j,p1cols[1]).value]) #Append its two parameters 
        try:
            if d[1]:
                c.append(d) #Once all the objects in a given image are found, move it to a master list
        except:
            pass
   
    return c

def pythag2(xb,xa,yb,ya): #the pythagorean theorem
    return (((xb-xa)**2)+((yb-ya)**2))**(0.5)

def compare2(list1,list2,title,dist=500):
    #Finds the nearest neighbor between 2 types of foci, aka the closest "list2" object to a "list1" object-
    #If the objects are further than max pixels in any direction (adjustable), or if there are none of the "list2" object
    #in any particular image, distance defaults to -1
    z=[]
    y=[]
    q=[]
    for i in range(len(list1)): #compare the image numbers in each data set
        z.append(list1[i][0])
    for j in range(len(list2)):
        y.append(list2[j][0])
    for i in z: #for all the images in the first set
        if i in y: #if the image also has objects in the second set
            for b in range(1,len(list1[z.index(i)])): #for each object...
                t=[]
                t.append(list1[z.index(i)][b]) #append the original coordinates...
                #now search for objects in the other category less than max distance away
                for c in range(1,len(list2[y.index(i)])): #for all objects
                    #filter so that you're only running the pythagorean theorem on objects
                    #within the default distance
                    if abs(list2[y.index(i)][c][0]-t[0][0])<dist: 
                        if abs(list2[y.index(i)][c][1]-t[0][1])<dist:
                            t.append(list2[y.index(i)][c]) #pass the close objects to a list

                if len(t)>1: #See if any objects were within range
                    
                    r=[dist*math.sqrt(2)] #Maximum that any value could achieve
                    for x in range(1,len(t)): #Run pythagorean theorem on all candidate objects
                        p=pythag2(t[0][0],t[x][0],t[0][1],t[x][1])
                        if p<r: #If it's the closest so far...
                            r=p #Replace the value 
                    q.append(r)        #append the final value to the list for writing
                else: #If no objects within range, append -1
                    q.append(-1)
        else: #If there were no objects in list 2, append -1
            for m in range(1,len(list1[z.index(i)])):
                q.append(-1)
    return q

def calcnearestneighbor(directory,csvoutname,par1="_Telom",par2="_Actin Foci", par3="_Filaments", par4="skjdflsj",
                        name1="TeloToActinFoci",name2="ActinFociToTelo",name3="TeloToActinFil",name4="ActinFilToTelo",
                        name5="Par1vs4", name6="Par4vs1",name7="ActinFociToFil",name8="ActinFilToFoci",name9="Par2vs4",
                        name10="Par4vs2",name11="Par3vs4",name12="Par4vs3",dist=500):
    #Defaults to finding telomeres and actin foci, but can compare up to 4 objects (par1, par2, par3, par4)-
    #make sure the first object is par1, second is par2 etc if you're only comparing 2 or 3 objects

    #If you want to change the defaults for compare objects (either the objects or the names), the "def" above is the place to do it

    q=combinecsv(directory,csvoutname) #Combines all the CSVs- can be done separately if needed
    rb=xlrd.open_workbook(q) #Open the excel file
    j=[]
    k=[]
    for sheet_name in rb.sheet_names():
        j.append(sheet_name)
        #Look for each category, if so append the index of the sheet that it's on
    for i in range(len(j)):
        if par1 in j[i]:
            k.append(i)
    for i in range(len(j)):
        if par2 in j[i]:
            k.append(i)
    for i in range(len(j)):
        if par3 in j[i]:
            k.append(i)
    for i in range(len(j)):
        if par4 in j[i]:
            k.append(i)
    wb=copy(rb) #Make the excel writable

    sobj1=rb.sheet_by_index(k[0]) #Use the sheet index from above to pull out the sheet with parameter 1
    ob1cols=sobj1.ncols #Figure out how many columns it has
    shobj1=wb.get_sheet(k[0]) #Now open the writable version of the sheet

    sobj2=rb.sheet_by_index(k[1])#Same thing with parameter 2
    ob2cols=sobj2.ncols
    shobj2=wb.get_sheet(k[1])
    
    a=pull2vals(q,k[0],firstvalue="Location_Center_X",secondvalue="Location_Center_Y") #pull the first location
    b=pull2vals(q,k[1],firstvalue="Location_Center_X",secondvalue="Location_Center_Y")  #pull the second location
    ab=compare2(a,b,name1,dist) #compare parameter 1 to parameter2
    shobj1.write(0,ob1cols,name1) #Write the title to the first open column
    for i in range(len(ab)): #Write all values
        shobj1.write(i+1,ob1cols,ab[i])
    ba=compare2(b,a,name2,dist) #compare par2 to par 1 (and so on)
    shobj2.write(0,ob2cols,name2)
    for i in range(len(ba)):
        shobj2.write(i+1,ob2cols,ba[i])

    if len(k)>2: #See if there is a 3rd object to be compared, if so open and compare, moving a column over if necessary
        ob3cols=sobj1.ncols
        shobj3=wb.get_sheet(k[2])
        c=pull2vals(q,k[2],firstvalue="Location_Center_X",secondvalue="Location_Center_Y")
        ac=compare2(a,c,name3,dist)
        shobj1.write(0,ob1cols+1,name3)
        for i in range(len(ac)):
            shobj1.write(i+1,ob1cols+1,ac[i])
        ca=compare2(c,a,name4,dist)
        shobj3.write(0,ob3cols,name4)
        for i in range(len(ca)):
            shobj3.write(i+1,ob3cols,ca[i])
        bc=compare2(b,c,name7,dist)
        shobj2.write(0,ob2cols+1,name7)
        for i in range(len(bc)):
            shobj2.write(i+1,ob2cols+1,bc[i])
        cb=compare2(c,b,name8,dist)
        shobj3.write(0,ob3cols+1,name8)
        for i in range(len(cb)):
            shobj3.write(i+1,ob3cols+1,cb[i])
    else:
        pass
    if len(k)>3: #4th is the same as the 3rd Object
        ob4cols=sobj1.ncols
        shobj4=wb.get_sheet(k[3])
        d=pull2vals(q,k[3],firstvalue="Location_Center_X",secondvalue="Location_Center_Y")
        ad=compare2(a,d,name5,dist)
        shobj1.write(0,ob1cols+2,name5)
        for i in range(len(ad)):
            shobj1.write(i+1,ob1cols+2,ad[i])
        da=compare2(d,a,name6,dist)
        shobj4.write(0,ob4cols,name6)
        for i in range(len(da)):
            shobj4.write(i+1,ob4cols,da[i])
        bd=compare2(b,d,name9,dist)
        shobj2.write(0,ob2cols+2,name9)
        for i in range(len(bd)):
            shobj2.write(i+1,ob2cols+2,bd[i])
        db=compare2(d,b,name10,dist)
        shobj4.write(0,ob4cols+1,name10)
        for i in range(len(db)):
            shobj4.write(i+1,ob4cols+1,db[i])
        cd=compare2(c,d,name11,dist)
        shobj3.write(0,ob3cols+2,name11)
        for i in range(len(cd)):
            shobj3.write(i+1,ob3cols+2,cd[i])
        dc=compare2(d,c,name12,dist)
        shobj4.write(0,ob4cols+2,name12)
        for i in range(len(dc)):
            shobj4.write(i+1,ob4cols+2,dc[i])
    else:
        pass
    r=os.path.join(csvoutname+'.xls') #Save the excel file
    wb.save(r)

if __name__=='__main__':

    direct=eg.diropenbox()
    csvout=eg.filesavebox(filetypes=["*.xls"])
    if not eg.ynbox('Do you want to temporarily change the defaults?', choices=('No','Yes')):
        msg='Enter new parameters-To permanently change defaults you must change the CSVCombiner sourcecode'
        title='Optional parameters'
        fieldNames=['Parameter 1 uniquely contains', 'Parameter 2 uniquely contains',
                        'Parameter 3 uniquely contains', 'Parameter 4 uniquely contains',
                        'Name of 1 vs 2','Name of 2 vs 1','Name of 1 vs 3','Name of 3 vs 1',
                        'Name of 1 vs 4','Name of 4 vs 1','Name of 2 vs 3','Name of 3 vs 2',
                        'Name of 2 vs 4','Name of 4 vs 2','Name of 3 vs 4','Name of 4 vs 3',
                        'Maximum distance(pixels)']
        fieldValues=["_Telom","_Actin Foci", "_Filaments", "skjdflsj","TeloToActinFoci",
                        "ActinFociToTelo","TeloToActinFil","ActinFilToTelo",
                        "Par1vs4", "Par4vs1","ActinFociToFil","ActinFilToFoci","Par2vs4",
                        "Par4vs2","Par3vs4","Par4vs3",500]
        a=eg.multenterbox(msg,title,fieldNames,fieldValues)
        calcnearestneighbor(direct,csvout,a[0],a[1],a[2],a[3],
                                       a[4],a[5],a[6],a[7],a[8],a[9],a[10],a[11],
                                       a[12],a[13],a[14],a[15],float(a[16]))
            
    else:
        calcnearestneighbor(direct,csvout)

def batchandremovetemps(direct):
    a=direct
    for i in os.listdir(a):
            for j in os.listdir(os.path.join(a,i)):
                    calcnearestneighbor(os.path.join(a,i,j),os.path.join(a,i)+j)
    for i in os.listdir(a):
        if 'temp' in i:
                    os.remove(os.path.join(a,i))
