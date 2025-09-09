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

st.set_page_config(layout="wide")

# config_path = 'config.json'
# with open(config_path, 'r') as config_file:
#     config = json.load(config_file)

# azure_api_key = config['azure_api_key']
# azure_api_version = config['azure_api_version']
# azure_endpoint = config['azure_endpoint']
# deployment_name = config['deployment_name']
# key = config['azure_cv_api_key']
# azure_cv_endpoint = config['azure_cv_endpoint']

# azure_document_api_key = config['azure_document_api_key']
# azure_document_endpoint = config['azure_document_endpoint']
# custom_model_id = config['custom_model_id']

azure_document_api_key = st.secrets["azure_document_api_key"]
azure_document_endpoint = st.secrets["azure_document_endpoint"]
custom_model_id = st.secrets["custom_model_id"]


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
        # st.write("Attempting prebuilt-layout analysis...")
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
        # st.write("✅ Layout analysis successful!")
        return result
    except Exception as e:
        st.error(f"Error processing invoice layout: {str(e)}")
        return None
def analyze_custom_model(uploaded_file):
    try:
        # st.write(f"🔍 Attempting custom model analysis")
        file_stream = BytesIO(uploaded_file.getvalue())
        if uploaded_file.name.lower().endswith('.pdf'):
            content_type = "application/pdf"
        elif uploaded_file.name.lower().endswith(('.jpg', '.jpeg')):
            content_type = "image/jpeg"
        elif uploaded_file.name.lower().endswith('.png'):
            content_type = "image/png"
        else:
            st.error("Unsupported file type. Please upload a PNG, JPG, JPEG, or PDF file.")
            return None
        
        poller = document_intelligence_client.begin_analyze_document(
            custom_model_id, 
            file_stream, 
            content_type=content_type
        )
        result = poller.result()
        # print("this is the custom model")
        # print(result)
        # st.write("Custom model analysis successful!")
        return result
    except Exception as e:
        st.error(f"Error processing custom model: {str(e)}")
        # st.write(f"Full error details: {type(e).__name__}: {e}")
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




def data_to_dataframe(invoice_data, custom_data=None):
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


    if custom_data:
        for doc in custom_data.documents:
            # print(f"🔍 Document has {len(doc.fields)} fields")
            for field_name, field in doc.fields.items():
                value = field.content if hasattr(field, 'content') and field.content else 'N/A'
                confidence = field.confidence if hasattr(field, 'confidence') else 'N/A'
                # print(f"🔍 Field: {field_name} = {value} (confidence: {confidence})")
                if value != 'N/A':
                    # Add custom fields only if they don't already exist
                    if field_name not in [data['Key'] for data in all_field_data]:
                        # print(f"✅ Adding new field: {field_name}")
                        all_field_data.append({'Key': field_name, 'Value': value, 'Confidence': confidence})

    fields_df = pd.DataFrame(all_field_data)
    # print(fields_df)
    tables_df = pd.concat(all_table_data, ignore_index=True) if all_table_data else pd.DataFrame()
    # print(tables_df)
    return fields_df, tables_df
 
def create_excel(fields_df, table_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
          current_row = 0
        
          if not fields_df.empty:
                worksheet = writer.book.create_sheet('Invoice_Data')
                worksheet.cell(row=current_row + 1, column=1).value = "=== INVOICE FIELDS ==="
                current_row += 2
                
                fields_df.to_excel(writer, sheet_name='Invoice_Data', startrow=current_row, index=False)
                current_row += len(fields_df) + 3  # Add some space
        
          if not table_df.empty:
                if fields_df.empty:
                    worksheet = writer.book.create_sheet('Invoice_Data')
                    current_row = 0
                
                worksheet.cell(row=current_row + 1, column=1).value = "=== INVOICE TABLES ==="
                current_row += 2
                
                table_df.to_excel(writer, sheet_name='Invoice_Data', startrow=current_row, index=False)
            
          if fields_df.empty and table_df.empty:
                empty_df = pd.DataFrame({'Message': ['No data extracted from the invoice']})
                empty_df.to_excel(writer, sheet_name='Invoice_Data', index=False)
            
          if 'Invoice_Data' in writer.sheets:
                worksheet = writer.sheets['Invoice_Data']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
          output.seek(0)
          return output

if 'fields_df' not in st.session_state:
    st.session_state.fields_df = pd.DataFrame()
if 'table_df' not in st.session_state:
    st.session_state.table_df = pd.DataFrame()
if 'ready_to_download' not in st.session_state:
    st.session_state.ready_to_download = False

# Streamlit app
st.title("Invoice Data Extraction")



uploaded_file = st.file_uploader("Upload your invoice PDF", type=["pdf","jpg", "png", "jpeg"])
if uploaded_file:
    # Extract data from PDF
    file_stream = BytesIO(uploaded_file.getvalue())
    st.write("Extracting data from the invoice...")



    if 'data_extracted' not in st.session_state:
        # invoice_data = analyze_invoice(uploaded_file)
        # custom_data = analyze_custom_model(uploaded_file)
        # layout_data = layout_invoice(uploaded_file)
        # st.session_state.data_extracted =True
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_text.text("Analyzing invoice structure...")
        invoice_data = analyze_invoice(uploaded_file)
        progress_bar.progress(33)
        status_text.text("Processing with custom model...")
        custom_data = analyze_custom_model(uploaded_file)
        progress_bar.progress(66)
        status_text.text("Extracting tables and layout...")
        layout_data = layout_invoice(uploaded_file)
        progress_bar.progress(100)
        
        status_text.text("Processing complete!")
        time.sleep(0.5)
        progress_bar.empty()
        status_text.empty()
        
        st.session_state.data_extracted = True

        if invoice_data and invoice_data.documents:
            fields_df, _ = data_to_dataframe(invoice_data, custom_data)
            st.session_state.fields_df = fields_df

        # if custom_data and custom_data.documents:
        #     for doc in custom_data.documents:
        #         for field_name, field in doc.fields.items():
        #             value = field.content if hasattr(field, 'content') and field.content else 'N/A'
        #             confidence = field.confidence if hasattr(field, 'confidence') else 'N/A'
        #             if value != 'N/A':
        #                 # st.write(f"**{field_name}**: {value} (Confidence: {confidence})")
        if layout_data:
            table_df = extract_table_data(layout_data)
            st.session_state.table_df = table_df
            # st.write(f"Table extraction: {'✅ Success' if not table_df.empty else '❌ No tables found'}")
    if not st.session_state.fields_df.empty:
        st.write("Extracted Field Data:")
        edited_fields = st.data_editor(st.session_state.fields_df, num_rows = "dynamic", key = "fields_editor", 
                                                    use_container_width=True)
        st.session_state.fields_df = edited_fields

    if not st.session_state.table_df.empty:
        st.write("Extracted Table data:")
        edited_tables = st.data_editor(st.session_state.table_df,num_rows="dynamic", key = "table_editor",
                                                   use_container_width=True)
        st.session_state.table_df = edited_tables

    if st.button('Finalize Edits'):
        st.session_state.ready_to_download = True
        st.success("Edits finalized. You can now download the Excel File")
    
    if st.button("Current Data Status"):
        st.write("**Current Fields Data:**")
        st.write(st.session_state.fields_df)
        st.write("**Current Table Data:**")
        st.write(st.session_state.table_df)
        # st.write(f"Ready to download: {st.session_state.ready_to_download}")

    if st.session_state.ready_to_download:
        # combined_df = pd.concat([st.session_state.fields_df ,st.session_state.table_df ],ignore_index=True)
        excel_file = create_excel(st.session_state.fields_df ,st.session_state.table_df)
        st.download_button(
            label="Download Excel file",
            data=excel_file,
            file_name="extracted_invoice_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload a PDF file to extract data")
    if 'data_extracted' in st.session_state:
        del st.session_state.data_extracted
    if 'fields_df' in st.session_state:
        st.session_state.fields_df = pd.DataFrame()
    if 'table_df' in st.session_state:
        st.session_state.table_df = pd.DataFrame()
    if 'ready_to_download' in st.session_state:
        st.session_state.ready_to_download = False






 