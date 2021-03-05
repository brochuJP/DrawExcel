# DrawExcel
Tool that draw VBA structure of you MS Excel file. 
It will scan VBA code in a MS Excel file and draw a visual diagram of its structure (Sheets, modules, forms ...) in an interractive SVG file.


## Requirement

It works only if you are on Windows OS with Microsoft Excel  


## Overview

This tool is primary to help people escape the Excel nightmare. It shows an overview of an Excel file structure. This aim to help developpers who need to maintain, update and document an Excel file that contain VBA.  


## How to generate the output

DrawExcel.DrawExcel( '--Your Excel file folder--' , '--Your excel file name (including extension)--')  


## Output Example:
<img alt="Overview of OEIS tools" src="docs/_MAIN.SVG" width="75%">

When you are working with a complex Excel file, you may find that for some VB components are chaotic. In order to have a better view of some section of your graph, click on the section name at the header of the sub zone to see a deeper view. 

For example if we click on Sheet3 we will see the following:

<img alt="Overview of OEIS tools" src="docs/Sheet3.SVG" width="75%">

## Dependencies


- All code is written in Python 3.
- Some code depends on:
  - pandas
  - win32com
  - graphviz




