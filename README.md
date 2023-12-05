# Fill_XLSX

Author: Miles Franklin<br>
Date:   12/04/2023<br>

### Description:
    This is the helpful template script to populate cells in an xlsx file
    based on the file names found in that file.

### Usages:
    python main.py --setup
        - Setup the XML files to show the demo. This will not work after 
        you begin to edit this script

### Demo:
    1. main.py read input.xlsx as input
    2. main.py itterates thorugh chunks of cells in the xlsx until a break is found
    3. an xml is parsed based on the "Source" field found in each block
    4. the subsequent cells are filled with content from the xml source
    5. output is save into output.xlsx 

### Files

input.xlsx<br>
<img width="174" alt="image" src="https://github.com/miles-franklin/Fill_XLSX/assets/101350365/afe49ab9-29d8-42ae-a682-b2c9a622d89e">

output.xlsx<br>
<img width="174" alt="image" src="https://github.com/miles-franklin/Fill_XLSX/assets/101350365/2f6ac4c4-f263-41d5-9135-0bd2364aab39">

inputs/a.xml, inputs/b.xml, inputs/c.xml<br>
<img width="415" alt="image" src="https://github.com/miles-franklin/Fill_XLSX/assets/101350365/f4329f4b-ac54-490c-85e2-9df9ab07b467">
