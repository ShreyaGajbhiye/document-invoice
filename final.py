import streamlit as st
from PIL import Image
import json
import pandas as pd
import io
import time
from io import BytesIO, StringIO
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
from msrest.authentication import CognitiveServicesCredentials
from openai import AzureOpenAI


# Load configuration

# config_path = 'config.json'
# with open(config_path, 'r') as config_file:
#     config = json.load(config_file)

azure_api_key = st.secrets["azure_api_key"]
azure_api_version = st.secrets["azure_api_version"]
azure_endpoint = st.secrets["azure_endpoint"]
deployment_name = st.secrets["deployment_name"]
key = st.secrets["azure_cv_api_key"]
azure_cv_endpoint = st.secrets["azure_cv_endpoint"]


azure_document_api_key = st.secrets["azure_document_api_key"]
azure_document_endpoint = st.secrets["azure_document_endpoint"]


model_id = 'prebuilt-invoice'
# Initialize Azure clients
document_analysis_client = DocumentAnalysisClient(
    endpoint=azure_document_endpoint,  
    credential=AzureKeyCredential(azure_document_api_key)
)

document_intelligence_client = DocumentIntelligenceClient(
    endpoint=azure_document_endpoint, 
    credential=AzureKeyCredential(azure_document_api_key))


def analyze_invoice(uploaded_file):
    try:
        file_stream = BytesIO(uploaded_file.getvalue())
        poller = document_analysis_client.begin_analyze_document("prebuilt-invoice", file_stream)
        result = poller.result()
        # print(result)
        return result
    except Exception as e:
        st.error(f"Error processing invoice: {str(e)}")
        return None
    
def layout_invoice(uploaded_file):
    try:
        file_stream = BytesIO(uploaded_file.getvalue())
        if uploaded_file.name.lower().endswith('.pdf'):
            content_type = "application/pdf"
        elif uploaded_file.name.lower().endswith(('.jpg','.jpeg')):
            content_type = "image/jpeg"
        elif uploaded_file.name.lower().endswith(('.png')):
            content_type = "image/png"
        else:
            st.error("Unsupported file type. Please upload a PNG, JPG, JPEG, or PNG file.")
            return None

        poller = document_intelligence_client.begin_analyze_document(
            "prebuilt-layout", 
            file_stream, 
            content_type=content_type)
        
        result = poller.result()
        # print(result)
        return result
    except Exception as e:
        st.error(f"Error processing invoice layout: {str(e)}")
        return None


def flatten_data(field, prefix=''):
    flat_data = {}
    if hasattr(field, 'value') and isinstance(field.value, dict):
        for sub_key, sub_field in field.value.items():
            flat_data.update(flatten_data(sub_field, prefix=f"{prefix}{sub_key}_"))
        print("flat_data_dict", flat_data)
    else:
        content = getattr(field, 'content', 'N/A') if hasattr(field, 'content') else 'N/A'
        flat_data[prefix.rstrip('_')] = content
    return flat_data

        
def extract_table_data(document):
    tables_list = []

    if hasattr(document, 'tables'):
        for table in document.tables:
            table_data = []
            for cell in table.cells:
                if len(table_data) <= cell.row_index:
                    table_data.extend([{} for _ in range(cell.row_index + 1 - len(table_data))])
                column_header = f"Column {cell.column_index}" 
                table_data[cell.row_index][column_header] = cell.content
            if table_data:
                df = pd.DataFrame(table_data)
                tables_list.append(df)

    if tables_list:
        return pd.concat(tables_list, ignore_index=True)
    else:
        return pd.DataFrame()




def data_to_dataframe(invoice_data):
    all_field_data = []
    all_table_data=[]

    for doc in invoice_data.documents:
        for field_name, field in doc.fields.items():
            value = field.content if hasattr(field, 'content') and field.content else 'N/A'
            confidence = field.confidence if hasattr(field, 'confidence') else 'N/A'
            all_field_data.append({'Key': field_name, 'Value': value, 'Confidence' : confidence})
    
        if hasattr(doc, 'tables'):
            table_data = extract_table_data(doc)
            if not table_data.empty:
                all_table_data.append(table_data)
    fields_df = pd.DataFrame(all_field_data)
    print(fields_df)
    tables_df = pd.concat(all_table_data, ignore_index=True) if all_table_data else pd.DataFrame()
    print(tables_df)
    return fields_df, tables_df



 
def create_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


# Initialize session states
if 'fields_df' not in st.session_state:
    st.session_state.fields_df = pd.DataFrame()
if 'table_df' not in st.session_state:
    st.session_state.table_df = pd.DataFrame()
if 'ready_to_download' not in st.session_state:
    st.session_state.ready_to_download = False

# Streamlit app setup
st.title("Invoice Data Extraction")

uploaded_file = st.file_uploader("Upload your invoice PDF", type=["pdf","jpg", "png", "jpeg"])
if uploaded_file:
    # Extract data from PDF
    file_stream = BytesIO(uploaded_file.getvalue())
    st.write("Extracting data from the invoice...")
    invoice_data = analyze_invoice(uploaded_file)
    layout_data = layout_invoice(uploaded_file)


    if invoice_data and invoice_data.documents:
        st.session_state.fields_df, _ = data_to_dataframe(invoice_data)
        if not st.session_state.fields_df.empty:
            st.write("Extracted Field Data:")
            edited_fields_df = st.data_editor(st.session_state.fields_df, num_rows = "dynamic")

    if layout_data:
        st.session_state.table_df = extract_table_data(layout_data)
        if not st.session_state.table_df.empty:
            st.write("Extracted Table data:")
            edited_tables_df = st.data_editor(st.session_state.table_df,num_rows="dynamic")

    if st.button('Finalize Edits'):
        st.session_state.ready_to_download = True
        st.success("Edits finalized. You can now download the Excel File")

    if st.session_state.ready_to_download:
        combined_df = pd.concat([st.session_state.fields_df ,st.session_state.table_df ],ignore_index=True)
        excel_file = create_excel(combined_df)
        st.download_button(
            label="Download Excel file",
            data=excel_file,
            file_name="extracted_invoice_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload a PDF file to extract data")






 
