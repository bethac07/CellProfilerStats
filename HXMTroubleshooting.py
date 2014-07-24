# -*- coding: utf-8 -*-
"""
Created on Tue Sep 24 04:25:10 2013

@author: Beth Cimini
"""

import HandyXLModules as HXM
import xlrd
inbook=xlrd.open_workbook(r'G:\TrackingOutputFolders200Frames\RPETelomereTrackingDif.xls')
cr=inbook.sheet_by_index(1)
tr=inbook.sheet_by_index(3)
crmsd=cr.col_values(10,1,201)
trmsd=tr.col_values(10,1,201)
HXM.graphmanyhists([crmsd,trmsd],['CRISPR','TRF1'],'blah','G:\trash2.png')
#HXM.graphmanyhists([crmsd,trmsd],['CRISPR','TRF1'],'blah','G:\trash2.png',normed=False)