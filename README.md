CellProfilerStats
=================

Programs for turning Cell Profiler data into graphs and meaningful statistics

This program takes .csv files output from Cell Profiler and allows you to add measurements between features, normalize intensities of children objects by those of their parent objects, and much more.
It also allows you to perform statistical analysis between treatments, and graphs the output while it's at it.
It isn't terribly customizable when run from the GUI, but the ability to create defaults allows you to pick the parameters you care about and look at them over and over again very quickly; great for optimization of your protocol of your analysis.
Working on building nice packaged distributions; in the meantime it requires python 2.7 with xlrd, xlwt, xlutils, easygui, matplotlib, statsmodels, PIL, and scipy.
