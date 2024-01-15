# SpAN
Span is a (very!) basic analysis tool for Laser Speckle imaging. It is written in VB6 and requires csv data files to be exported from the moor FLPI analysis software. 

# Using the software 
The software works with CSV files exported from moorFLPI V5 Review software. The format of the file is as follows, with the following headers. Note, at the moment, SpAN can not discriminate between multiple ROIs, so each file should only contain the data from a single ROI. 


       moorFLPI V5.0 Image Statistics (Example.mmf)
    ==============================================================
    
    ROI No	Frame No	Flux Mean	Flux Std	Flux Med	Flux Min	Flux Max	DC Mean	DC Std	DC Med	DC Min	DC Max	Total Pixel	Valid Pixel	Valid %	Total Area	Valid Area	Frame Time
    2	1	706.7	253.0	661	247	3216	151.1	11.4	152	123	174	6081	6081	100.0	2.2	2.2	00:00:00.000
    2	2	660.9	196.8	628	281	2443	150.4	11.5	151	122	176	6081	6081	100.0	2.2	2.2	00:00:00.040
    2	3	807.7	207.6	780	362	2376	150.6	12.0	151	122	172	6081	6081	100.0	2.2	2.2	00:00:00.080
    2	4	814.3	196.9	786	367	1965	151.4	12.8	152	122	176	6081	6081	100.0	2.2	2.2	00:00:00.120
    2	5	993.7	231.8	956	529	2368	152.7	12.7	152	124	179	6081	6081	100.0	2.2	2.2	00:00:00.160
    2	6	1003.5	206.1	970	565	2004	153.3	12.5	153	125	176	6081	6081	100.0	2.2	2.2	00:00:00.200
    2	7	880.4	170.4	855	483	1676	151.5	12.3	151	122	176	6081	6081	100.0	2.2	2.2	00:00:00.240
    2	8	762.3	147.7	741	419	1502	150.5	12.1	151	122	176	6081	6081	100.0	2.2	2.2	00:00:00.280
    2	9	785.7	165.3	762	423	2463	150.0	12.1	150	122	173	6081	6081	100.0	2.2	2.2	00:00:00.320
    2	10	831.2	170.5	811	429	1971	151.6	12.7	152	122	176	6081	6081	100.0	2.2	2.2	00:00:00.360
