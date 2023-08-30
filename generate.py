from tkinter.filedialog import askopenfile, askdirectory
from functions import *
import pandas as pd

class SAP:
    def __init__(self, sheet, name, output_path):
        self.sheet = sheet
        self.name = name
        self.compiler = self.compiler()
    
    def compiler(self):
        print('Generating: ', self.name)
        self.input_unit = input_reader(self.sheet)
        self.output_data = match_xml(self.input_unit)
        if self.output_data:
            self.writer = self.writer()
        else:
            raise Exception('No output data')
    
    def writer(self):
        xml_string = data_to_xml(self.output_data)
        file_path = output_path+'\{}.xml'.format(self.name)
        with open(file_path,'w') as f:
            try:
                f.write(prettify(xml_string))
                findAndReplace(file_path)
            except:
                raise Exception('Invalid XML structure')
        pass

if __name__ == "__main__":
    excel_path = askopenfile(title='Select SAP Calc Sheet', filetypes=[("Excel files", ".xlsx .xls")])
    output_path = askdirectory()
    print('SAP Calc Sheet: '+excel_path.name)
    df = pd.read_excel(excel_path.name, header=1, sheet_name=None)

    for name, sheet in df.items():
        if 'Unit' in name:
            sap = SAP(sheet,name,output_path)