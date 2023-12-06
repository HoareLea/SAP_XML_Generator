"""Holds Class for generating the output"""

import warnings
import pandas as pd
import os
import streamlit as st
from functions import input_reader, match_xml, data_to_xml, prettify, find_and_replace, generate_unique_name, delete_files_in_directory

warnings.simplefilter(action='ignore', category=UserWarning)

class SAP:
    """The main Class for handling the inputs and outputs"""
    def __init__(self, sheet, name):
        self.sheet = sheet
        self.name = name
        self.input_unit = input_reader(self.sheet)
        self.output_data = match_xml(self.input_unit)
        if self.output_data:
            self.writer()
            self.remove_temp_files()
        else:
            raise st.error('No output data')

    def writer(self):
        """Writes out the xml data to a file"""
        xml_string = data_to_xml(self.output_data)
        temp_file_name = generate_unique_name(self.name)
        
        self.temp_path = './temp/'
        if not os.path.exists(self.temp_path):
            os.makedirs(self.temp_path)

        file_path = f'{self.temp_path}/{temp_file_name}.xml'
        with open(file_path,'w', encoding="utf-8") as f:
            try:
                f.write(prettify(xml_string))
                self.xml_out = find_and_replace(file_path)
            except Exception as exc:
                raise st.error('Invalid XML structure') from exc
        pass

    def remove_temp_files(self):
        """Removes all temp files"""
        delete_files_in_directory(self.temp_path)
        pass

def generate(file):
    excel_path = file
    print('SAP Calc Sheet: '+excel_path.name)
    df = pd.read_excel(excel_path, header=1, sheet_name=None)
    units_xmls = []
    units_names = []
    for this_name, this_sheet in df.items():
        if 'Unit' in this_name:
            sap = SAP(this_sheet,this_name)
            units_xmls.append(sap.xml_out)
            units_names.append(sap.name)
    return units_names,units_xmls
