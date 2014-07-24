# -*- coding: utf-8 -*-
"""
Created on Tue Apr 03 10:49:23 2012

@author: Beth Cimini
"""

from uncertainties import ufloat
from uncertainties.umath import *

def checkprop(m1,s1,m2,s2):
    x=ufloat((m1,s1))
    y=ufloat((m2,s2))
    print y/x
    
checkprop(0.000358,0.000163,0.000384,0.000172)