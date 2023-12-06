import streamlit as st
from io import BytesIO
import zipfile
from generate import generate
import base64
import datetime

def main():
    st.title("SAP XML Generator")
    st.header('Download the standard Excel Calc Sheet', divider='rainbow')
    
    standard_calc_sheet = 'CALC-XX-XX-SAP CALC TEMPLATE.xlsx'
    if st.button('Download Excel File'):
        with open(standard_calc_sheet, 'rb') as file:
            file_content = file.read()
            bin_str = base64.b64encode(file_content).decode()
        st.markdown(f'<a href="data:application/octet-stream;base64,{bin_str}" download="{standard_calc_sheet}">Download standard Excel Calc Sheet</a>', unsafe_allow_html=True)
    
    st.header('Generate the XMLs by uploading your completed Calc Sheet below', divider='rainbow')
    
    # File upload widget
    uploaded_file = st.file_uploader("Choose a Calc Sheet file", type=["xlsx"])

    if uploaded_file is not None:
        # Process file and generate XML
        if st.button("Generate XML"):
            with st.spinner("Processing..."):
                names,xml_outputs = generate(uploaded_file)

            # Create a zip file containing all the XML files
            timestamp = datetime.datetime.now()
            zip_filename = f"{str(timestamp).split('.')[0].replace(':','-')}_SAP_XMLs.zip"
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for i, xml_output in enumerate(xml_outputs):
                    xml_bytes = xml_output.encode('utf-8')
                    zip_file.writestr(f'{names[i]}.xml', xml_bytes)
           
            st.markdown(
                f'<a href="data:application/zip;base64,{base64.b64encode(zip_buffer.getvalue()).decode()}" download="{zip_filename}">Download {zip_filename}</a>',
                unsafe_allow_html=True
            )
            
if __name__ == "__main__":
    main()