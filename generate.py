"""Holds Class for generating the output"""

from tkinter.filedialog import askopenfile, askdirectory
import warnings
import pandas as pd
from functions import input_reader, match_xml, data_to_xml, prettify, find_and_replace, ErrorFound

warnings.simplefilter(action='ignore', category=UserWarning)

class SAP:
    """The main Class for handling the inputs and outputs"""
    def __init__(self, sheet, name, output_path):
        self.sheet = sheet
        self.name = name
        self.input_unit = input_reader(self.sheet)
        self.output_data = match_xml(self.input_unit)
        self.compiler(output_path)

    def compiler(self, output_path):
        """Begins the writing process if there is data to outputs"""
        print('Generating: ', self.name)
        if self.output_data:
            self.writer(output_path)
        else:
            raise ErrorFound('No output data')

    def writer(self, output_path):
        """Writes out the xml data to a file"""
        xml_string = data_to_xml(self.output_data)
        file_path = output_path+f'\{self.name}.xml'
        with open(file_path,'w', encoding="utf-8") as f:
            try:
                f.write(prettify(xml_string))
                find_and_replace(file_path)
            except Exception as exc:
                raise ErrorFound('Invalid XML structure') from exc
        pass

if __name__ == "__main__":
    excel_path = askopenfile(
        title='Select SAP Calc Sheet',
        filetypes=[("Excel files", ".xlsx .xls")]
    )
    this_output_path = askdirectory()
    print('SAP Calc Sheet: '+excel_path.name)
    df = pd.read_excel(excel_path.name, header=1, sheet_name=None)

    for this_name, this_sheet in df.items():
        if 'Unit' in this_name:
            sap = SAP(this_sheet,this_name,this_output_path)
