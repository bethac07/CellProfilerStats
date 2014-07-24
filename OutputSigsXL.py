import xlrd
import numpy
from scipy import stats
import xlwt
from HandyXLModules import *
import easygui as eg
import shelve
import os
from matplotlib.backends.backend_pdf import PdfPages
from statsmodels.sandbox.stats.multicomp import multipletests


titles= xlwt.easyxf('font: name Times New Roman;'
'pattern: pattern no_fill;'
'borders: left thin, right thin, top thin, bottom thin;'
'alignment: wrap true')
othercells=xlwt.easyxf('pattern: pattern no_fill;'
'borders: left thin, right thin, top thin, bottom thin')
vsmall = xlwt.easyxf('font: name Times New Roman;'
'borders: left thin, right thin, top thin, bottom thin;' 
'pattern: fore_colour light_blue, pattern solid_fill;'
'alignment: wrap true')
small = xlwt.easyxf('font: name Times New Roman;'
'borders: left thin, right thin, top thin, bottom thin;'
'pattern: fore_colour pale_blue, pattern solid_fill;'
'alignment: wrap true')
big = xlwt.easyxf('font: name Times New Roman;'
'borders: left thin, right thin, top thin, bottom thin;'
'pattern: fore_colour rose, pattern solid_fill;'
'alignment: wrap true')
vbig = xlwt.easyxf('font: name Times New Roman;'
'borders: left thin, right thin, top thin, bottom thin;'
'pattern: fore_colour red,pattern solid_fill;'
'alignment: wrap true')
hist=xlwt.easyxf('font: height 4000')

def pickafilter(typeofthing,param,questiontoask,unfiltvalues,parentinfo=None):
    filterkeys=[]
    filtervals=[]
    intentered=False
    howtofilt=eg.indexbox('How do you want to filter '+questiontoask+'?',choices=('By a numerical value', 'By a percentile'))
    if howtofilt==0: #if they say by numverical value
        while intentered==False:
            filt=eg.multenterbox(msg='How do you want to filter '+questiontoask+'?',fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))#let the user input the filter they want
            try:
                isint=int(filt[1])
                filterkeys.append(param+':'+filt[0]+filt[1])
                filtervals.append(unfiltvalues[:2]+[typeofthing,filt])
                intentered=True
            except:
                intentered=False
    else:
        toporbottom=eg.indexbox(msg='Do you want to look above or below a specified percentile of' +param+'?',choices=('Above','Below','One set of each'))
        if toporbottom==0:
            while intentered==False:
                toppercent=eg.enterbox(msg='Enter the number for the percentile of' +param+' you want to look above (ie at the top ____ percent')
                try:
                    isint=int(toppercent)
                    toappendtos=['Percentile>=',toppercent]
                    filterkeys.append(param+':'+toappendtos[0]+toappendtos[1])
                    filtervals.append(unfiltvalues[:2]+[typeofthing,toappendtos])
                    intentered=True
                except:
                    intentered=False
        elif toporbottom==1:
            while intentered==False:
                botpercent=eg.enterbox(msg='Enter the number for the percentile of ' +param+' you want to look below (ie at the bottom ____ percent')
                try:
                    isint=int(botpercent)
                    toappendtos=['Percentile<=',botpercent]
                    filterkeys.append(param+':'+toappendtos[0]+toappendtos[1])
                    filtervals.append(unfiltvalues[:2]+[typeofthing,toappendtos])
                    intentered=True
                except:
                    intentered=False
        else:
            intenteredsecond=False
            while intentered==False:
                toppercent=eg.enterbox(msg='Enter the number for the percentile of ' +param+' you want to look above (ie at the top ____ percent')
                try:
                    isint=int(toppercent)
                    toappendtos=['Percentile>=',toppercent]
                    filterkeys.append(param+':'+toappendtos[0]+toappendtos[1])
                    filtervals.append(unfiltvalues[:2]+[typeofthing,toappendtos])
                    intentered=True
                except:
                    intentered=False
            while intenteredsecond==False:
                botpercent=eg.enterbox(msg='Enter the number for the percentile of ' +param+' you want to look below (ie at the bottom ____ percent')
                try:
                    isint=int(botpercent)
                    toappendtos=['Percentile<=',botpercent]
                    filterkeys.append(param+':'+toappendtos[0]+toappendtos[1])
                    filtervals.append(unfiltvalues[:2]+[typeofthing,toappendtos])
                    intenteredsecond=True
                except:
                    intenteredsecond=False
    #print filterkeys,filtervals
    return filterkeys,filtervals

def choosewhichstatssin(book):
    pulldefs=shelve.open(os.path.join(os.curdir,'OutputSigsXLshelf'),writeback=True)
    
    if eg.ynbox(msg='Do you want to use the defaults?'):
        whichdefault=eg.choicebox('Which default set do you want to use?',choices=pulldefs.keys())
        g=pulldefs[whichdefault]

    else:
        #Pull all the possible parameters
        t=[]
        colsbysheet=[]
        a=readsheets1file(book)
        for i in a:
            if i != 'Image': #generally unhelpful output from cell profiler (has 150+ columns)- this could be removed or more sheets could be added
                ii=str('%.2d' %a.index(i))
                c=colheadingreadernum(book,a.index(i))
                newsheet=[]
                for j in c:
                    jj=str('%.3d' %c.index(j))
                    t.append(ii+':)'+jj+':)'+str(i)+'-'+str(j))
                    newsheet.append(ii+':)'+jj+':)'+str(i)+'-'+str(j))
                    colsbysheet.append(newsheet)
        
        xparams=[]
        paramdict={}
        fx=eg.multchoicebox(msg='Select parameters',choices=t)
        addmore=True
        while addmore==True:
            eg.textbox(msg='So far you have selected the following parameters. On the next screen you will be asked if you want to add any more.',text='\n'.join(fx))
            moretodo=eg.ynbox('Do you need to add any more parameters?')
            if moretodo==0:
                addmore=False
            else:
                gx=eg.multchoicebox(msg='Which of these previously unselected parameters do you want to add?', choices=list(set(t)-set(fx)))
                fx+=gx
                addmore=True
        for param in range(len(fx)):
            paramdict[fx[param][9:]]=[int(fx[param][0:2]),int(fx[param][4:7])]
        donefiltering=False
        while donefiltering==False:
            filtf=eg.multchoicebox(msg='Do you want to filter any of these?',choices=(paramdict.keys()))
            for param in paramdict.keys():
                if len(paramdict[param])==2:
                    paramdict[param].append(0)
                if param in filtf:
                    heads=colheadingreadernum(book,paramdict[param][0])
                    numorstr=eg.indexbox(('How do you want to filter '+param),choices=['By a specific value or percentile','By another measure of that same object','By relationship to another object'])
                    if numorstr==0: #if they say by numverical value
                        filtkeys,filtvals=pickafilter(1,param,param,paramdict[param])
                        for waystofilt in range(len(filtkeys)):
                            paramdict[filtkeys[waystofilt]]=filtvals[waystofilt]
                    elif numorstr==1:
                        whichfilt=eg.multchoicebox('Which other measure(s) do you want to sort '+param+' by?',choices=heads)
                        for eachfilt in whichfilt:
                            filtkeys,filtvals=pickafilter(2,param+': by '+eachfilt,eachfilt+' in breaking down '+param,paramdict[param])
                            #eachfiltsort=eg.multenterbox(msg='What filter do you want to put on '+str(eachfilt)+'?',fields=('Operator- choose from ==, !=, <,>, <=,>=','Value'))
                            for waystofilt in range(len(filtkeys)):
                                paramdict[filtkeys[waystofilt]]=filtvals[waystofilt][:3]+[[[paramdict[param][0],heads.index(eachfilt),filtvals[waystofilt][3]]]]                           
                    elif numorstr==2:
                        relatives=[]
                        relativesindex=[]
                        for eachhead in range(len(heads)):
                            if 'Children' in heads[eachhead]:
                                relativesindex.append(eachhead)
                                relatives.append(heads[eachhead][:heads[eachhead].index('_',9)])
                            elif 'Parent' in heads[eachhead]:
                                relativesindex.append(eachhead)
                                relatives.append(heads[eachhead])
                        whichrels=eg.multchoicebox('Which related objects do you want to sort '+param+' by?',choices=relatives)
                        for rels in whichrels:
                            if 'Children' in rels:
                                relname=rels[9:]
                            else:
                                relname=rels[7:]
                            relsheet=a.index(relname)
                            relheads=colheadingreadernum(book,relsheet)
                            if 'Children' in rels:
                                myparentcolumn=relheads.index('Parent_'+param[:param.index('-')])
                            whichrelthings=eg.multchoicebox('Which parameter of '+relname+' do you want to sort '+param+' by?',choices=relheads)
                            for relthings in whichrelthings:
                                filtkeys,filtvals=pickafilter(2,param+': by '+rels+ '-'+relthings,relthings+' of '+relname+' as a way to break down'+param,paramdict[param])
                                if 'Parent' in rels:
                                    for waystofilt in range(len(filtkeys)):
                                        paramdict[filtkeys[waystofilt]]=filtvals[waystofilt][:3]+[[[paramdict[param][0],relativesindex[relatives.index(rels)],relsheet,relheads.index(relthings),filtvals[waystofilt][3]]]]
                                else:
                                    for waystofilt in range(len(filtkeys)):
                                        paramdict[filtkeys[waystofilt]]=filtvals[waystofilt][:3]+[[[paramdict[param][0],relativesindex[relatives.index(rels)],relsheet,myparentcolumn,relheads.index(relthings),filtvals[waystofilt][3]]]]
        
            allparams=paramdict.keys()
            allparams.sort()
            eg.textbox(msg='So far you have generated the following parameters. On the next screen you will be asked if you want to add more filters.',text='\n'.join(allparams))
            if eg.ynbox('Do you want to add any other filters?'):
                donefiltering=False
            else:
                donefiltering=True
        removefromparams=eg.multchoicebox('Do you want to remove any parameters from this list? This will be your last chance to do so.',choices=paramdict.keys())
        for toremove in removefromparams:
            paramdict.pop(toremove)
        graphaslog=eg.multchoicebox('Do you want to graph any of the following in log scale (alone or in addition to linear scale)?',choices=paramdict.keys())
        if graphaslog!=[]:
            graphonlylog=eg.multchoicebox('Which of these do you want to graph in log scale ONLY? (Unselected items will be graphed as linear and log)',choices=graphaslog )
        else:
            graphonlylog=[]
        for param in paramdict:
            if param in graphonlylog:
                paramdict[param].append(2)
            elif param in graphaslog:
                paramdict[param].append(1)
            else:
                paramdict[param].append(0)
            xparams.append(paramdict[param])
        xparams.sort()
        g=[xparams] #return the list of parameters the user wants to compare
        if eg.ynbox('Do you want to save these settings as a new default?'):
            newdefname=eg.enterbox(msg='Give this default a descriptive identifier')
            pulldefs[newdefname]=g
        
    pulldefs.close()
    #print g
    return g

def findtheps(list1):
    basemed=float(numpy.median(list1[0]))   
    changelist=[]
    mwplist=[]
    ksplist=[]


    for i in range(1,len(list1)):
            if basemed!=0:
                changelist.append(100*(float(numpy.median(list1[i]))-basemed)/basemed)
            else:
                changelist.append('n.d.-> base median=0')
            #print list1[0],list1[i]
            try:
                mwplist.append(stats.mannwhitneyu(list1[0],list1[i])[1])
            except ValueError:
                mwplist.append(1)
                print 'Raised Error'
            ksplist.append(stats.ks_2samp(list1[0],list1[i])[1])
    return changelist,mwplist,ksplist
    #print mwplist,ksplist
    
def untangleolddefaults(olddefaultlist):
    output=[]
    inputlist=olddefaultlist[0]
    for i in inputlist:
        if i[-1]==0:
            output.append(i)
        elif len(i[-1])==1:
            output.append(i)
        else:
            for el in i[-1]:
                output.append(i[:3]+[[el]])
    toreturn=[output]
    return toreturn

def runeachone(writebook,names,paths,whichfiles,runcount,outdir,outfilename):
    histdic={}
    listofstats=[]
    graphlin=[]
    graphlog=[]
    if len(whichfiles)==1:
        ident=outfilename[:-4]
        while len(ident)>31:
            ident=eg.enterbox(ident+' is too long to be a sheet name- enter a shorter name')
        basebook=xlrd.open_workbook(paths[names.index(whichfiles[0])])
        if whichfiles[0][-7:-4]=='Out':
            readablelabellist=[whichfiles[0][:-7]]
        else:
            readablelabellist=[whichfiles[0][:-4]]
        whichstats=choosewhichstatssin(basebook)
        #print whichstats
        basevals=arrangedivsin(basebook,whichstats[0])
        if len(basevals)!=len(whichstats[0]):
            whichstats=untangleolddefaults(whichstats)
            olddefault=True
        else:
            olddefault=False
        for i in range(len(basevals)):
            #print basevals[i][0]
            #print i, whichstats[0][i],basevals[i]
            if len(basevals[i])>1:
                listofstats.append(basevals[i][0])
                histdic[basevals[i][0]]=[basevals[i][1:]]
                if olddefault==False: #exclude old defaults with a filter
                    if whichstats[0][i][-1]==2:
                        graphlog.append(basevals[i][0])
                    elif whichstats[0][i][-1]==1:
                        graphlin.append(basevals[i][0])
                        graphlog.append(basevals[i][0])
                    else:
                        graphlin.append(basevals[i][0])   
                else:
                    graphlin.append(basevals[i][0])
                writesheet.write(1+i,0, basevals[i][0],titles)
    else:
        baseline=eg.choicebox(msg='Which file is the baseline?', choices=whichfiles)
        ident=baseline[:-4]+' as baseline'
        while len(ident)>31:
            ident=eg.enterbox(ident+' is too long to be a sheet name- enter a shorter name')
        whichfilescopy=[]
        for filename in whichfiles:
            whichfilescopy.append(filename)
        whichfilescopy.remove(baseline)
        labelist=[]
        labelist.append(baseline)
        labelist=labelist+whichfilescopy
        readablelabellist=[]
        for i in labelist:
            if i[-7:-4]=='Out':
                readablelabellist.append(i[:-7])
            else:
                readablelabellist.append(i[:-4])
        basebook=xlrd.open_workbook(paths[names.index(baseline)])
        compbook=xlrd.open_workbook(paths[names.index(whichfilescopy[0])])
        whichstats=choosewhichstatssin(basebook)
        basevals=arrangedivsin(basebook,whichstats[0])
        writesheet=writebook.add_sheet(ident)
        writesheet.col(0).width=9000
        for others in range(len(whichfilescopy)):
            writesheet.col(others+1).width=(14400/len(whichfilescopy))
            writesheet.col(others+1).set_style(othercells)
        writesheet.write(0,0,ident,titles)
        #print whichstats[0]
        #print len(basevals),len(whichstats[0])
        if len(basevals)!=len(whichstats[0]):
            whichstats=untangleolddefaults(whichstats)
            olddefault=True
        else:
            olddefault=False
        for i in range(len(basevals)):
            #print basevals[i][0]
            #print i, whichstats[0][i],basevals[i]
            if len(basevals[i])>1:
                listofstats.append(basevals[i][0])
                histdic[basevals[i][0]]=[basevals[i][1:]]
                if olddefault==False: #exclude old defaults with a filter
                    if whichstats[0][i][-1]==2:
                        graphlog.append(basevals[i][0])
                    elif whichstats[0][i][-1]==1:
                        graphlin.append(basevals[i][0])
                        graphlog.append(basevals[i][0])
                    else:
                        graphlin.append(basevals[i][0])   
                else:
                    graphlin.append(basevals[i][0])
                writesheet.write(1+i,0, basevals[i][0],titles)
            else:
                print basevals[i][0]," no objects met that filter"
        for i in range(len(whichfilescopy)):
            writesheet.write(0,i+1,whichfilescopy[i],titles)
            compbook=xlrd.open_workbook(paths[names.index(whichfilescopy[i])])
            compvals=arrangedivsin(compbook,whichstats[0])
            for m in range(len(compvals)):
                if compvals[m][0] in histdic.keys():
                    histdic[compvals[m][0]].append(compvals[m][1:])
        mannwhitforcorr=[]
        ksforcorr=[]
        percentchangesnested=[]
        for i in range(len(listofstats)):
            #print listofstats[i]
            percentchange,mannwhitney,ks=findtheps(histdic[listofstats[i]])
            percentchangesnested.append(percentchange)
            mannwhitforcorr+=mannwhitney
            ksforcorr+=ks
        #print multipletests(mannwhitforcorr,method='h')
        try:
            mwbools,mwvals,z,y=multipletests(mannwhitforcorr,method='h')
        except ValueError:
            mwbools=len(mannwhitforcorr)*[True]
            mwvals=len(mannwhitforcorr)*[1.000]
        try:
            ksbools,ksvals,z,y=multipletests(ksforcorr,method='h')
        except ValueError:
            ksbools=len(ksforcorr)*[True]
            ksvals=len(ksforcorr)*[1.000]
        itercount=0
        for parameter in range(len(percentchangesnested)):
            #print sigtest
            for treatment in range(len(percentchangesnested[parameter])):
                change=percentchangesnested[parameter][treatment]
                if type(change)==str:
                    towrite=', '+change
                    changeneg=0
                else:
                    if change<0:
                        changeneg=True
                    else:
                        changeneg=False
                    towrite=', '+"%0.2f" %change +'% change in median'
                if mwbools[itercount]==True:
                    if ksbools[itercount]==True:
                        if changeneg==True:
                            writesheet.write(parameter+1,treatment+1,'M.W. <'+"%0.3f" %mwvals[itercount]+', K.S. <'+"%0.3f" %ksvals[itercount]+towrite,vsmall)
                        elif changeneg==False:
                            writesheet.write(parameter+1,treatment+1,'M.W. <'+"%0.3f" %mwvals[itercount]+', K.S. <'+"%0.3f" %ksvals[itercount]+towrite,vbig)
                        else:
                            writesheet.write(parameter+1,treatment+1,'M.W. <'+"%0.3f" %mwvals[itercount]+', K.S. <'+"%0.3f" %ksvals[itercount]+towrite,titles)
                    else:
                        if changeneg==True:
                            writesheet.write(parameter+1,treatment+1,'M.W. <'+"%0.3f" %mwvals[itercount]+towrite,small)
                        elif changeneg==False:
                            writesheet.write(parameter+1,treatment+1,'M.W. <'+"%0.3f" %mwvals[itercount]+towrite,big)
                        else:
                            writesheet.write(parameter+1,treatment+1,'M.W. <'+"%0.3f" %mwvals[itercount]+towrite,titles)
                else:
                    if ksbools[itercount]==True:
                        if changeneg==True:
                            writesheet.write(parameter+1,treatment+1,'K.S. <'+"%0.3f" %ksvals[itercount]+towrite,small)
                        elif changeneg==False:
                            writesheet.write(parameter+1,treatment+1,'K.S. <'+"%0.3f" %ksvals[itercount]+towrite,big)
                        else:
                            writesheet.write(parameter+1,treatment+1,'K.S. <'+"%0.3f" %ksvals[itercount]+towrite,titles)
                    else:
                        writesheet.write(parameter+1,treatment+1,'n.s.'+towrite,titles)
                itercount+=1
                """if len(sigtest[j])<7:
                    writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],titles)
                elif len(sigtest[j])<10:
                    if sigtest[j][-1]=='-':
                        writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],small)
                    elif sigtest[j][-1]=='+':
                        writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],big)
                    else:
                        writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],titles)
                else:
                    if sigtest[j][-1]=='-':
                        writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],vsmall)
                    elif sigtest[j][-1]=='+':
                        writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],vbig)
                    else:
                        writesheet.write(i+1,j+1,sigtest[j]+sigvals[j],titles)"""
    writesheet.flush_row_data()
    if runcount==0:
        if eg.ynbox('Do you want to run the histograms?'):
            histPDF=PdfPages(os.path.join(outdir,outfilename[:-4]+'Histograms.pdf'))
            graphshelf=shelve.open(os.path.join(outdir,outfilename[:-4]+'Histshelf'),writeback=True)
            histsheet=writebook.add_sheet('Histograms')
            for i in range(2):
                histsheet.col(i).width=16500
            initrowcount=((len(graphlin)+len(graphlog)))
            for i in range(initrowcount):
                histsheet.row(i).set_style(hist)
            saveas=os.path.join(outdir,'trash')
            enrichlist=[]
            histcount=0
            for stattouse in listofstats:
                graphlinbool=False
                graphlogbool=False
                if stattouse in graphlin:
                    graphlinbool=True
                    graphshelf[str(histcount)]=['graphswarm',histdic[stattouse],readablelabellist,saveas,stattouse]
                    graphswarm(histdic[stattouse],readablelabellist,saveas,stattouse,size=(450,330),PDF=histPDF)
                    histsheet.insert_bitmap(saveas+'.bmp',histcount/2,0)
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                    histcount+=1
                    graphshelf[str(histcount)]=['graphscumhist',histdic[stattouse],readablelabellist,saveas,stattouse]
                    graphscumhist(histdic[stattouse],readablelabellist,saveas,stattouse,size=(450,330),PDF=histPDF)
                    histsheet.insert_bitmap(saveas+'.bmp',histcount/2,1)
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                    histcount+=1
                if stattouse in graphlog:
                    graphlogbool=True
                    graphshelf[str(histcount)]=['graphswarm',histdic[stattouse],readablelabellist,saveas,stattouse,'log=True']
                    graphswarm(histdic[stattouse],readablelabellist,saveas,stattouse,size=(450,330),PDF=histPDF,log=True)
                    histsheet.insert_bitmap(saveas+'.bmp',histcount/2,0)
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                    histcount+=1
                    graphshelf[str(histcount)]=['graphscumhist',histdic[stattouse],readablelabellist,saveas,stattouse,'log=True']
                    graphscumhist(histdic[stattouse],readablelabellist,saveas,stattouse,size=(450,330),PDF=histPDF,log=True)
                    histsheet.insert_bitmap(saveas+'.bmp',histcount/2,1)
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                    histcount+=1
                if '(' in stattouse:
                    findunfiltered=stattouse[:stattouse.index(r'(')]
                    if findunfiltered in listofstats:
                        histdic[findunfiltered+'outgroup of'+stattouse]=[]
                        for treatgroup in range(len(histdic[stattouse])):
                            wholegroup=histdic[findunfiltered][treatgroup]
                            ingroup=histdic[stattouse][treatgroup]
                            for ingroupval in ingroup:
                                if ingroupval in wholegroup:
                                    wholegroup.remove(ingroupval)
                            histdic[findunfiltered+'outgroup of'+stattouse].append(wholegroup)
                        if graphlinbool==True:
                            enrichlist.append((histdic[stattouse],histdic[findunfiltered+'outgroup of'+stattouse],readablelabellist,saveas,'Relative value of '+stattouse,False))
                        if graphlogbool==True:
                            enrichlist.append((histdic[stattouse],histdic[findunfiltered+'outgroup of'+stattouse],readablelabellist,saveas,'Relative value of '+stattouse,True))
                            
            for i in range((len(enrichlist)+1)/2):
                histsheet.row(initrowcount+i).set_style(hist)
            for i in range(len(enrichlist)):
                enr=enrichlist[i]
                graphshelf[str(histcount)]=['graphsubgroupswarm',enr[0],enr[1],enr[2],enr[3],enr[4],'log=enr[5]']
                graphsubgroupswarm(enr[0],enr[1],enr[2],enr[3],enr[4],size=(450,330),PDF=histPDF,log=enr[5])
                histsheet.insert_bitmap(saveas+'.bmp',initrowcount+(i/2),i%2)
                os.remove(saveas+'.bmp')
                os.remove(saveas+'.png')

            histPDF.close()
            graphshelf.close()
            histsheet.flush_row_data()
        if eg.ynbox('Do you want to run any scatter plots on the parameters just analyzed?'):
            scattPDF=PdfPages(os.path.join(outdir,outfilename[:-4]+'Scatters.pdf'))
            scattshelf=shelve.open(os.path.join(outdir,outfilename[:-4]+'Scattshelf'))
            if eg.ynbox('Do you want to save the resulting hi-res vector images? (Warning, there will be a LOT of files)'):
                savescatts=True
                subdir=os.path.join(outdir,'scatterimages')
                if not os.path.isdir(subdir):
                    os.mkdir(subdir)
                saveas=os.path.join(subdir,'trash')
            else:
                savescatts=False
                saveas=os.path.join(outdir,'trash')
            scattdefs=shelve.open(os.path.join(os.curdir,'OSXLscatter'),writeback=True)
            if eg.ynbox('Do you want to use a default?'):
                whichdefault=eg.choicebox('Which default set do you want to use?',choices=scattdefs.keys())
                scattstorun=scattdefs[whichdefault][0]
                bubblestorun=scattdefs[whichdefault][1]
            else:
                scattstorun=[[],[]]
                bubblestorun=[[],[],[]]
                readabletorun=[]
                finished=False
                while finished==False:
                    itemforx=eg.choicebox('Which statistic do you want to have as the x-axis',choices=listofstats)
                    subitems=[]
                    dashindex=itemforx.index('-')
                    for i in listofstats:
                        if itemforx[:dashindex]==i[:dashindex]:
                            if itemforx[dashindex:]!=i[dashindex:]:
                                subitems.append(i)
                    itemsfory=eg.multchoicebox('Which statistic(s) do you want to have on the y axis against '+itemforx+'?',choices=subitems)
                    itemsforbubbles=eg.multchoicebox('Which statistic(s) vs '+itemforx+' do you want to also add a 3rd parameter/bubblegraph?',choices=itemsfory)
                    for i in range(len(itemsfory)):
                        scattstorun[0].append(itemforx)
                        scattstorun[1].append(itemsfory[i])
                        readabletorun.append(itemforx+' vs '+itemsfory[i])
                        if itemsfory[i] in itemsforbubbles:
                            tobebubbled=eg.multchoicebox('Which statistic(s) do you want to see as marker size against '+itemforx+' vs '+itemsfory[i]+'?',choices=subitems)
                            for r in tobebubbled:
                                bubblestorun[0].append(itemforx)
                                bubblestorun[1].append(itemsfory[i])
                                bubblestorun[2].append(r)
                                readabletorun.append(itemforx+' vs '+itemsfory[i]+' vs '+r)
                    eg.textbox(msg='So far you have selected',text='\n'.join(readabletorun))
                    if eg.ynbox(msg='Are there any others you would like to add?'):
                        finished=False
                    
                    else:
                        if eg.ynbox('Do you want to save these settings as a new default?'):
                            newdefname=eg.enterbox(msg='Give this default a descriptive identifier')
                            scattdefs[newdefname]=[scattstorun,bubblestorun]
                        finished=True
            scattdefs.close()
            scattsheet=writebook.add_sheet('Scatterplots')
            for i in range(2):
                scattsheet.col(i).width=16500
            for i in range(((len(scattstorun[0])+len(bubblestorun[0]))/2)+1):
                scattsheet.row(i).set_style(hist)
            
            count=0
            for i in range(len(scattstorun[0])):
                #print scattstorun[0][i],scattstorun[1][i]
                scattshelf[str(count)]=['graphscatter',histdic[scattstorun[0][i]],histdic[scattstorun[1][i]],readablelabellist,saveas,scattstorun[0][i],scattstorun[1][i],'savefiles=savescatts','log=x']
                graphscatter(histdic[scattstorun[0][i]],histdic[scattstorun[1][i]],readablelabellist,saveas,scattstorun[0][i],scattstorun[1][i],size=(450,330),savefiles=savescatts,PDF=scattPDF,log='x')
                scattsheet.insert_bitmap(saveas+'.bmp',count/2,count%2)
                count+=1
                os.remove(saveas+'.bmp')
                os.remove(saveas+'.png')
            for i in range(len(bubblestorun[0])):
                scattshelf[str(count)]=['graphbubble',histdic[bubblestorun[0][i]],histdic[bubblestorun[1][i]],histdic[bubblestorun[2][i]],readablelabellist,saveas,bubblestorun[0][i],bubblestorun[1][i],bubblestorun[2][i], 'savefiles=savescatts']
                graphbubble(histdic[bubblestorun[0][i]],histdic[bubblestorun[1][i]],histdic[bubblestorun[2][i]],readablelabellist,saveas,bubblestorun[0][i],bubblestorun[1][i],bubblestorun[2][i], size=(450,330),savefiles=savescatts,PDF=scattPDF)            
                scattsheet.insert_bitmap(saveas+'.bmp',count/2,count%2)
                count+=1
                os.remove(saveas+'.bmp')
                os.remove(saveas+'.png')
            scattPDF.close()
            scattshelf.close()
            for k in range(len(labelist)):
                scattPDF=PdfPages(os.path.join(outdir,ident+'-'+labelist[k][:-4]+'Scatters.pdf'))
                scattsheet=writebook.add_sheet(labelist[k])
                print labelist[k]
                for i in range(2):
                    scattsheet.col(i).width=16500
                for i in range(((len(scattstorun[0])+len(bubblestorun[0]))/2)+1):
                    scattsheet.row(i).set_style(hist)
                count=0
                for i in range(len(scattstorun[0])):
                    #print scattstorun[0][i],scattstorun[1][i]
                    graphscatter([histdic[scattstorun[0][i]][k]],[histdic[scattstorun[1][i]][k]],[readablelabellist[k]],saveas,scattstorun[0][i],scattstorun[1][i],size=(450,330),savefiles=savescatts,PDF=scattPDF,log='x')
                    scattsheet.insert_bitmap(saveas+'.bmp',count/2,count%2)
                    count+=1
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                for i in range(len(bubblestorun[0])):
                    graphbubble([histdic[bubblestorun[0][i]][k]],[histdic[bubblestorun[1][i]][k]],[histdic[bubblestorun[2][i]][k]],[readablelabellist[k]],saveas,bubblestorun[0][i],bubblestorun[1][i],bubblestorun[2][i], size=(450,330),savefiles=savescatts,PDF=scattPDF)            
                    scattsheet.insert_bitmap(saveas+'.bmp',count/2,count%2)
                    count+=1
                    os.remove(saveas+'.bmp')
                    os.remove(saveas+'.png')
                scattPDF.close()
            
        
               
        

def dothestuff(direct=0):
    if direct==0:
        direct=eg.diropenbox()
    a=findexcel(direct)
    b=[]
    c=[]
    for i in a:
        b.append(i[0])
        c.append(i[1])
    whichfiles=eg.multchoicebox(msg='Which input files do you want to use?', choices=b)

    w=eg.filesavebox(msg='What do you want to name the output file?',filetypes=["*.xls"])+'.xls'
    writebook=xlwt.Workbook() #make a new file
    outdir,filename=os.path.split(w)
    
    runcount=0
    another=True
    while another==True:
        runeachone(writebook,b,c,whichfiles,runcount,outdir,filename)
        runcount+=1
        if not eg.ynbox('Do you want to run the same input and output files but with a different baseline?'):
            another=False
    writebook.save(w)


if __name__=='__main__':
    dothestuff()
