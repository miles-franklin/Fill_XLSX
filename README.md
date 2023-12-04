# Fill_XLSX

Author: Miles Franklin
Date:   12/04/2023
Description:
    This is the helpful template script to populate cells in an xlsx file
    based on the file names found in that file.

Usages:
    python main.py --setup
        - Setup the XML files to show the demo. This will not work after 
        you begin to edit this script

Demo:
    1. main.py read input.xlsx as input
    2. main.py itterates thorugh chunks of cells in the xlsx until a break is found
    3. an xml is parsed based on the "Source" field found in each block
    4. the subsequent cells are filled with content from the xml source
    5. output is save into output.xlsx 