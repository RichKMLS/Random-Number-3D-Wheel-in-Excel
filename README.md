# Random-Number-3D-Wheel-in-Excel :green_book:
This project demonstrates the creative use of VBA to animate 3D shapes in Excel. I created a 3D wheel with 12 segments; the wheel spins and slows down to randomly stop at a selected segment, which flashes green to indicate the winner. The spinning duration varies each spin, adding suspense and excitement to the animation. This project showcases my skills and creativity in using Excel beyond its conventional functions.

![theamazingspinningwheel](https://user-images.githubusercontent.com/105183376/172466157-d9ce2d85-3eac-4029-b77e-136451a1179c.png)

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
                   
                   Average:	42	
        Standard Deviation:	8	
	
