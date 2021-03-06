# DrawExcel
Tool that draws VBA structure of your MS Excel file. 
It will scan VBA code in a MS Excel file and draws a visual diagram of its structure (Sheets, modules, forms ...) in an interractive SVG file.


## Requirement

It works only if you are on Windows OS with Microsoft Excel  


## Overview

This tool is primary to help people escape the Excel nightmare. It shows an overview of an Excel file structure. This aim to help developpers who need to maintain, update and document an Excel file that contain VBA.  


## How to generate the output

DrawExcel.DrawExcel( '--Your Excel file folder--' , '--Your excel file name (including extension)--')  


## Output Example:
![](https://github.com/brochuJP/DrawExcel/blob/main/docs/_MAIN.jpg?raw=true)

When you are working with a complex Excel file, you may find that some VB components are chaotic. In order to have a better view of some section of your graph, click on the section name at the header of the sub zone to see a deeper view. 

For example if we click on Sheet3 we will see the following:

![](https://github.com/brochuJP/DrawExcel/blob/main/docs/Sheet3.jpg?raw=true)
  
By doing this you can can have a deeper view of you graph.
## Dependencies


- All code is written in Python 3.
- Some code depends on:
  - pandas
  - win32com
  - graphviz




