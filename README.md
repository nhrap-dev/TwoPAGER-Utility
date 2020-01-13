# TwoPAGER

!!Marked for update!!

* If you have Anaconda installed, the GUI should launch from double clicking the .bat file and no config modification is necessary.

* If you do not have Anaconda installed, you can try opening a terminal in the ./src folder and running `python TwoPAGER_Utility.py`. If you do not have the correct libraries installed, this might not work.


!!Old documentation below!!:

Prepares the files from a Hazus earthquake run to be send to the USGS for creating the TwoPAGER report.

<h2>To use: </h2>
Only modify the Config.ini file

The file reads as
```
[VARIABLES]
scenario_name: ak_regionM7
folder_path: H:\Events\National\Alaska_2018
ftp: False
```

1) change ```scenario_name``` to your scenario.
2) change ```folder_path``` to the folder path you want the output stored.

<h3> An example change would look like </h3>

``` 
[VARIABLES]
scenario_name: Napa_geomean
folder_path: C:\Disasters\Napa
ftp: False 
```
