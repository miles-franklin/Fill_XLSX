"""
    Author: Miles Franklin
    Date:   12/04/2023
    Description:
        This is the helpful template script to populate cells in an xlsx file
        based on the file names found in that file.
    
    Usages:
        python main.py --setup
            - Setup the XML files to show the demo. This will not work after 
            you begin to edit this script
"""

#===============================================================================
# IMPORTS
#===============================================================================
import os
import sys
import random
import argparse
import openpyxl
import xml.etree.ElementTree as ET

#===============================================================================
# FUNTIONS
#===============================================================================
def get_args(args):
    parser = argparse.ArgumentParser(description='')
    parser.add_argument("-s", "--setup", action="store_true",
                        help='Creates XML files if needed for demo.')
    return parser.parse_args(args)

def setup_xmls():
    inputs_dir = os.path.dirname(os.path.abspath(__file__))
    inputs_dir = os.path.join(inputs_dir, "inputs")
    if not os.path.exists(inputs_dir):
        os.makedirs(inputs_dir)

    for xml_file_name in ["a.xml", "b.xml", "c.xml"]:
        # Create the root element
        root = ET.Element("worksheet", xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        # Create the data element
        data = ET.SubElement(root, "data")

        # Create var_1, var_2, and var_3 elements and add them to the data element
        ET.SubElement(data, 'var_1').text = str(random.randint(1000, 9999))
        ET.SubElement(data, 'var_2').text = str(random.randint(1000, 9999))
        ET.SubElement(data, 'var_3').text = str(random.randint(1000, 9999))

        # Create the ElementTree object
        tree = ET.ElementTree(root)
        ET.indent(tree)

        # Write the XML content to a file
        tree.write("output.xml", encoding="utf-8", xml_declaration=True)

        # Write the tree to an XML file
        print(f"Creating {os.path.join(inputs_dir, xml_file_name)}")
        if os.path.exists(os.path.join(inputs_dir, xml_file_name)):
            os.remove(os.path.join(inputs_dir, xml_file_name))
        tree.write(file_or_filename=os.path.join(inputs_dir, xml_file_name), encoding="utf-8", xml_declaration=True)

def get_data_from_xml(source="inputs/a.xml", namespace=""):
    tree = ET.parse(source)
    root = tree.getroot()
    ouput = {}

    if namespace == "":
        namespace = root.tag.split('}')[0] + "}"

    for child in root.find(f'{namespace}data'):
        ouput[child.tag.replace(namespace, "")] = child.text
    
    return ouput

#===============================================================================
# MAIN
#===============================================================================
def main():
    print(sys.argv[1:])
    args = get_args(sys.argv[1:])  

    if args.setup:
        setup_xmls()

    # Load the workbook
    workbook = openpyxl.load_workbook('input.xlsx')

    # Select the active worksheet
    worksheet = workbook.active

    cell_a1 = worksheet["A1"].value
    rows_offset = 5
    i = 0
    while True:
        base_location = worksheet["A1"].offset(row=i*rows_offset, column=0).coordinate
        if worksheet[base_location].value != "Source":
            break

        # Get input from xml
        source_loc = worksheet[base_location].offset(row=0, column=1).coordinate
        data = get_data_from_xml(source=worksheet[source_loc].value)
        
        # Write from xml to xlsx
        for i in range(3):
            var_loc = worksheet[base_location].offset(row=i+1, column=0).coordinate
            input_loc = worksheet[base_location].offset(row=i+1, column=1).coordinate

            # Write to input_loc
            worksheet[input_loc] = data[worksheet[var_loc].value]
        
        # Increment
        i += 1

    # Save the workbook
    workbook.save('output.xlsx')

if __name__ == "__main__":
    main()