# Random-Number-3D-Wheel-in-Excel :green_book:
Used to demonstrate how VBA can be integrated with 3D shapes to deliver smooth and simple animations.
![theamazingspinningwheel](https://user-images.githubusercontent.com/105183376/172466157-d9ce2d85-3eac-4029-b77e-136451a1179c.png)
In this example—a number will be randomly selected from a wheel labeled 1 through 12. It may choose this number very quickly or it may draw out the process through the use of 'suspense' decided by RNG.
----
###### Build Instructions

        1. Import ThisWorkbook.cls and randomWheel.bas into theAmazingSpinningWheel.xlsx
        2. If you wish to archive your results: set archiveDir in randomWheel.bas to the directory you wish to save ArchiveResults.csv
        3. Right click on the wheel and assign macro to either 'archiveruns' or 'spinWheel'
        4. Save file as theAmazingSpinningWheel.xlsm and close file.
----        
###### For best results run the following via command prompt or through a shortcut:

        "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" /e /x "C:\Path\to\theAmazingSpinningWheel.xlsm"

----
###### The following is data retrieved using archivereader.bas after performing 500 archived runs.
      
        Values      1	2	3	4	5	6	7	8	9	10	11	12
	
        Totals	50	42	29	47	49	35	52	45	53	32	31	35
        N.Dist	2.90%	4.72%	1.53%	3.87%	3.24%	3.46%	2.23%	4.37%	1.92%	2.45%	2.13%	3.46%

                   
                   Average:	42	
        Standard Deviation:	8	
	
