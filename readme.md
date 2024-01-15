# SpAN
Span is a (very!) basic analysis tool for Laser Speckle imaging. It is written in VB6 and requires csv data files to be exported from the moor FLPI analysis software. Before loading in the first data file, click on Update in the first tab area to save the default settings for the software to memory. Recommended values:

       Low:                1550
       Hi:                 2550
       XStep:              10
       GridY:              100
       GridX:              25
       Grid Font Size:     14

# Using the software 
The software works with CSV files exported from moorFLPI V5 Review software. The format of the file is as follows, with the following headers. Note, at the moment, SpAN can not discriminate between multiple ROIs, so each file should only contain the data from a single ROI. 


       moorFLPI V5.0 Image Statistics (Example.mmf)
    ==============================================================
    
    ROI No	Frame No	Flux Mean	Flux Std	Flux Med	Flux Min	Flux Max	DC Mean	DC Std	       DC Med	       DC Min	       DC Max	       Total Pixel	Valid Pixel	Valid %	Total Area	Valid Area	Frame Time
    2	       1	       706.7	       253.0	       661	       247	       3216	       151.1	       11.4	       152	       123	       174	       6081	       6081	       100.0	       2.2	       2.2	       00:00:00.000
    2	       2	       660.9	       196.8	       628	       281	       2443	       150.4	       11.5	       151	       122	       176	       6081	       6081	       100.0	       2.2	       2.2	       00:00:00.040
    2	       3	       807.7	       207.6	       780	       362	       2376	       150.6	       12.0	       151	       122	       172	       6081	       6081	       100.0	       2.2	       2.2	       00:00:00.080
    2	       4	       814.3	       196.9	       786	       367	       1965	       151.4	       12.8	       152	       122	       176	       6081	       6081	       100.0	       2.2	       2.2	       00:00:00.120
    ....

Data files should be loaded in under the Data file option on the program window. Once loaded, the graph will populate with the data from the file. You do not need to remove any of the additional information from the CSV file (header etc), these will be automatically disregarded. 

There are a number of options available on the left hand side:

## Highlight inversions
