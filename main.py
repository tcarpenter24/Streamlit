import streamlit as st
import os
import tempfile
import zipfile
from docx import Document
from PyPDF2 import PdfReader
import openpyxl
import subprocess
import sys
import streamlit.web.cli as stcli

# Keywords data for functional requirements
keywords_data = {
    "Group Accounts": {
        "main_topics": ["access control", "account management", "User matrix", "roles and responsibilities"],
        "detail_keywords": ["group account"],
        "Information Needed": "Are group accounts used?",
        "Analyst Input": "",
        "Extra Details": "",
        "Artifact Reference": "",
        "Corresponding CCIs": "CCI-002129, CCI-002140, CCI-002141, CCI-002142",
    },
    "Temporary Accounts": {
        "main_topics": ["roles and responsibilities", "access control", "account management", "User matrix"],
        "detail_keywords": ["temporary account"],
        "Information Needed": "Are temporary accounts used?",
        "Analyst Input": "",
        "Extra Details": "",
        "Artifact Reference": "",
        "Corresponding CCIs": "CCI-000016, CCI-001361",
    },
    "Contingency Training": {
        "main_topics": ["contingency plan", "emergency policy", "contingency", "training plan"],
        "detail_keywords": ["contingency training", "training", "exercise"],
        "Information Needed": "Is there contingency training?",
        "Analyst Input": "",
        "Extra Details": "",
        "Artifact Reference": "",
        "Corresponding CCIs": "CCI-000486, CCI-000485, CCI-000487, CCI-000488",
    }
}

# Fetch keywords based on the requirement
def fetch_keywords(requirement):
    return keywords_data[requirement]

# Read file content and search for main topics and detail keywords
def read_file_content(filepath, main_topics, detail_keywords):
    main_topic_matches = {}
    detail_keyword_matches = {}
    if filepath.endswith('.pdf'):
        reader = PdfReader(filepath)
        for page_num, page in enumerate(reader.pages, start=1):
            text = page.extract_text()
            for keyword in main_topics:
                if keyword.lower() in text.lower():
                    main_topic_matches.setdefault(keyword, []).append(f"Page {page_num}")
            for keyword in detail_keywords:
                if keyword.lower() in text.lower():
                    detail_keyword_matches.setdefault(keyword, []).append(f"Page {page_num}")
    elif filepath.endswith('.txt'):
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            for keyword in main_topics:
                if keyword.lower() in content.lower():
                    main_topic_matches.setdefault(keyword, ['In document'])
            for keyword in detail_keywords:
                if keyword.lower() in content.lower():
                    detail_keyword_matches.setdefault(keyword, ['In document'])
    elif filepath.endswith('.docx'):
        doc = Document(filepath)
        for i, para in enumerate(doc.paragraphs, start=1):
            for keyword in main_topics:
                if keyword.lower() in para.text.lower():
                    main_topic_matches.setdefault(keyword, []).append(f"Paragraph {i}")
            for keyword in detail_keywords:
                if keyword.lower() in para.text.lower():
                    detail_keyword_matches.setdefault(keyword, []).append(f"Paragraph {i}")
    return main_topic_matches, detail_keyword_matches

# Extract and search files for matches
def extract_and_search(zip_file_path, control):
    keywords = fetch_keywords(control)
    main_topics = keywords['main_topics']
    detail_keywords = keywords['detail_keywords']
    detailed_matches = {}
    with tempfile.TemporaryDirectory() as tempdir:
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            files_to_extract = [f for f in zip_ref.namelist() if f.endswith(('.docx', '.txt', '.pdf')) and not os.path.basename(f).startswith('~$')]
            zip_ref.extractall(tempdir, members=files_to_extract)
        for file_path in files_to_extract:
            local_file_path = os.path.join(tempdir, file_path)
            if os.path.isfile(local_file_path):
                main_topic_matches, detail_keyword_matches = read_file_content(local_file_path, main_topics, detail_keywords)
                if main_topic_matches or detail_keyword_matches:
                    detailed_matches[local_file_path] = {
                        "main_topics": main_topic_matches,
                        "detail_keywords": detail_keyword_matches
                    }
    return detailed_matches, keywords['Information Needed'], keywords['Analyst Input'], keywords['Corresponding CCIs']

# Save uploaded file to a temporary directory
def save_uploaded_file(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        return tmp_file.name

# Update Excel file with the findings
def update_excel(data, excel_path):
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook['CCI Report']
        
        ccis_to_find = [cci.strip() for cci in data['Corresponding CCIs'].split(",")]
        updated_rows = set()

        for row in range(1, sheet.max_row + 1):
            cell_value = str(sheet[f'D{row}'].value).strip()
            if cell_value in ccis_to_find:
                sheet[f'J{row}'] = data['Finding Details']
                updated_rows.add(row)

        workbook.save(excel_path)
        workbook.close()

        if updated_rows:
            updated_rows_str = ", ".join(str(r) for r in sorted(updated_rows))
            st.success(f"Excel file successfully updated! Rows updated: {updated_rows_str}")
        else:
            st.warning("No matching CCIs found to update in the Excel file.")
        
    except PermissionError as e:
        st.error(f"Failed to update Excel: Permission denied. {e}")
    except Exception as e:
        st.error(f"An error occurred while updating Excel: {e}")

# Streamlit main application
def main_app():
    st.title("Cyber Survey Tool")
    uploaded_zip = st.file_uploader("Upload ZIP file containing documents", type=['zip'])
    requirement = st.selectbox("Select Functional Requirement", list(keywords_data.keys()))
    excel_path = r"C:\Cyber Survey Tool\Test_Assessment_General_RMF_Export.xlsx"

    if uploaded_zip and st.button("Search Documents"):
        temp_file_path = save_uploaded_file(uploaded_zip)
        search_results, info_needed, analyst_input, ccis = extract_and_search(temp_file_path, requirement)
        st.session_state.search_results = search_results
        st.session_state.info_needed = info_needed
        st.session_state.analyst_input = analyst_input
        st.session_state.ccis = ccis

    if 'checked_documents' not in st.session_state:
        st.session_state.checked_documents = []

    if 'search_results' in st.session_state and st.session_state.search_results:
        for doc_path, matches in st.session_state.search_results.items():
            main_topics_matches = matches["main_topics"]
            detail_keyword_matches = matches["detail_keywords"]
            if detail_keyword_matches:
                doc_name = os.path.basename(doc_path)
                st.write(f"**{doc_name}** found keywords at: ")
                for k, v in detail_keyword_matches.items():
                    st.write(f"  - Detail keyword '{k}': {', '.join(v)}")
                checkbox_key = f"select_{doc_name}_{hash(doc_path)}"
                if st.checkbox(f"Select {doc_name}", key=checkbox_key):
                    if doc_name not in st.session_state.checked_documents:
                        st.session_state.checked_documents.append(doc_name)
                else:
                    if doc_name in st.session_state.checked_documents:
                        st.session_state.checked_documents.remove(doc_name)
        with st.form(key='details_form'):
            st.write("**Information Needed:**", st.session_state.info_needed)
            analyst_input = st.radio("Analyst Input", ["Yes", "No"])
            extra_details = st.text_area("Extra Details")
            artifact_reference = st.text_input("Artifact Reference", value=", ".join(st.session_state.checked_documents))
            st.write("**Corresponding CCIs:**", st.session_state.ccis)
            submitted = st.form_submit_button("Submit Details")

            if submitted:
                finding_details = f"{requirement} {'are' if analyst_input == 'Yes' else 'are not'} used.\n{extra_details}\nArtifacts: {artifact_reference}"
                export_data = {
                    "Functional Requirement": requirement,
                    "Finding Details": finding_details,
                    "Corresponding CCIs": st.session_state.ccis,
                    "Artifact Reference": artifact_reference
                }
                st.session_state.export_data = export_data
                st.success("Details submitted successfully! You can now proceed or submit additional information.")
                update_excel(export_data, excel_path)

def run():
    # Determine the path to the main.py script
    streamlit_script = os.path.join(os.path.dirname(__file__), "main.py")
    
    # Check if the script exists
    if not os.path.exists(streamlit_script):
        print(f"Error: {streamlit_script} does not exist")
        sys.exit(1)
    
    # Check if the script is being run as the main module
    if os.getenv('RUN_MAIN', False):
        # Run the streamlit script with subprocess
        subprocess.run([sys.executable, "-m", "streamlit", "run", streamlit_script, "--server.port", "8501"])
    else:
        # Set an environment variable to indicate the script is being run
        os.environ['RUN_MAIN'] = 'true'
        # Exit and run the Streamlit CLI main function
        sys.exit(stcli.main())

if __name__ == '__main__':
    main_app()