CellProfilerStats
=================

Programs for turning Cell Profiler data into graphs and meaningful statistics

This program takes .csv files output from Cell Profiler and allows you to add measurements between features, normalize intensities of children objects by those of their parent objects, calculate ratios between numbers of objects or between different wavelengths, and much more.

It also allows you to perform statistical analysis between different datasets (ie different treatmeants, genotypes, different versions of the same protocol for optimization , and graphs the output while it's at it.  Right now this is set for two non-parametric statistical tests (Mann-Whitney and KS, both corrected by Holm-Bonferroni), with swarm plots and cumulative frequency plots of all data analyzed returned as vector graphs in a PDF.  Optimization for 96- and 384- well plates is in progress.

It is moderately customizable when run from the GUI, and the ability to create defaults allows you to pick the parameters you care about and look at them over and over again very quickly; great for optimization of your protocol or of your Cell Profiler analysis.  

Working on building nice packaged distributions; in the meantime it requires python 2.7 with xlrd, xlwt, xlutils, easygui, matplotlib, statsmodels, PIL, and scipy.

Blanket disclaimer- outside of one 10-session class, I have no formal training in programming.  I built all of this for personal use- I'm thrilled it's worked as well as it has and that it's becoming useful for others.  Therefore, expect it may be buggy and unpolished, and while I'll get to fixes as soon as I can, this is a tool and a side project for me so be reasonable in amount of time it takes to get things done.
