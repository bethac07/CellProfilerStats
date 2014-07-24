import easygui as eg
import sys


while 1:
    message='What do you want to do?'
    title='Cell Profiler Statistics Wizard'
    choices=('Process CSV files','Combine Excel files', 'Generate graphs and statistics', 
            'Work with tracking files','Exit')
    mainmenu=eg.indexbox(message,title,choices)

    if mainmenu==0:
        import CSVCombinerGeneral as CSVC
        if not eg.ynbox('Do you want to do the EXACT same operations on multiple data sets (ie use batch mode)?'):
            direct=eg.diropenbox('Which folder has your output files in it?')
            csvout=eg.filesavebox('Where do you want to save your output file?',filetypes=["*.xls"])
            CSVC.calcnearestneighbor(direct,csvout)
        else:
            CSVC.batchmode()
            
    if mainmenu==1:
        import XLCombiner as XLC
        XLC.runxlcombiner()

    if mainmenu==2:
        import OutputSigsXL as OutSXL
        OutSXL.dothestuff()
            
    if mainmenu==3:
        firstorsecond=eg.boolbox(msg='Do you want to process a .csv file or run summary statistics?',
                                 choices=['Process .csv', 'Run summary statistics'])
        if firstorsecond:
            import SimpleTrackingCleanup as STC
            STC.dothestuff()
        else:
            import TTStats as TTS
            TTS.dothestuff()
    
    if mainmenu==4:
        if eg.ccbox('Exit?','Exit?',choices=('No','Yes')):
            pass
        else:
            sys.exit(0)

