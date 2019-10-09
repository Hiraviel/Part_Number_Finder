<h2 align="center">Excel Part Number Finder (Or text finder)</h2>

<b>Tkinter GUI to search an especific text within an Excel file.</b>

This Script requires the installation of this library:

    pip install xlrd
    
The GUI needs two Inputs:

    Number Part: Or whatever the word you need to find inside an Excel file.
    Directory: Where the Excel files are located.
    
After that, just click the button "Buscar" and the script will attempt to find the "Number Part".

If succeeded, or not, 3 things going to happen:

    1.- A pop-up window will show the total of elements found within the Excel Files. Just press "Aceptar"
    2.- The file/s that had the "Number Part" will be listed at the bottom of the GUI.
        You can select it, copy and paste at your Windows Explorer or throught the "Run Command".
    3.- A "Number_Part".txt file will be generated, inside the folder where the script where launched.
        This .txt file will contain the list described in point 2.

<h4>Notes:</h4>

    A.- This Script only works with ".xlsx" and ".xls" extensions due to limitations of the "xlrd" library.
    B.- You need to CLOSE the Preview Panel, if not closed the Script will crash.
    C.- You can paste the "Number Part" in the GUI, but cannot copy it from the GUI :C
    D.- You can paste the Directory but with "/" instead of "\", and the same as point C cannot copy it from the GUI.
    E.- You need to remove all the temp files within the folder, corresponding to the Excel files.
        Inside the Script there is a function that can handle it, but where not implemented.
    F.- When executed the Script appears not to work, but it is executing its function.
    
