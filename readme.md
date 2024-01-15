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

## Additional options

### Highlight inversions
Inflection points will be labelled on the graph. Down points (i.e. low inflection points) will be marked in Cyan, whilst Up points (i.e. high infection points) will be labelled in green. Disregarded inflections will be marked in white. Disregarded inflections occur when the point is invalid; i.e. the inflection is both not preceeded by two points concurrently in the same direction and not followed by two points concurrently in the opposite direction.

### Markers
You can draw both point markers for data points and slice numbers onto the graph. Enable these options to see the points. 

### XStep
How many pixels should appear between each point in the X axis. The smaller the number, the closer the points will be together on the graph.

## Show (Up/Down points)
The box to the right hand side of this option lists all of the inflection points either in the upward or downward direction, depending on the setting in this drop down menu. The values in the listbox can be copied to the clipboard using the button underneath, and then pasted into Excel or any other analysis software.

## Video subsampling (experimental)
SpAN can use subsampling to create a video based on the up or down points as selected above. It requires that the video frames match up directly with the frame number in the CSV file. Once the video is loaded, the following actions take place: 
* The video file is turned into individual image frames
* The frames that correspond with the up and down points are selected from the video frames and used to generate a final video of just those points
* The resulting video is output to the desktop.
