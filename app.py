import io
from datetime import date
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

import streamlit as st

from src.service import GenerationService

st.set_page_config(page_title="CPR Invoice Generator", layout="wide")
st.title("CPR Invoice and Letter Generator")
st.write("Upload an Excel file, enter invoice details, and download the generated output as a ZIP file.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
invoice_no = st.text_input("Invoice Number")
course_options = [
    "Preparatory Cyber Crime",
    "Eow & Cyber Crime",
    "Use of AI",
    "Women Safety",
    "SC/ST Act",
    "UAPA Act",
    "Crime Against Women",
    "Cyber Crime Investigation",
    "Bharosa Cell",
    "ATS",
    "Woman Safety",
    "Gender Sensitization & Forensic Training Programme for Police Personnel",
    "Training of Trainers",
    "EOW and Cyber Crime",
    "Preparatory Course on Cyber Crime",
]
course_name = st.selectbox("Course Name", course_options)
col1, col2 = st.columns(2)
with col1:
    training_from = st.date_input("Training From Date", value=date.today())
with col2:
    training_to = st.date_input("Training To Date", value=date.today())
col3, col4 = st.columns(2)
with col3:
    invoice_date = st.date_input("Invoice Date", value=date.today())
with col4:
    letter_date = st.date_input("Letter Date", value=date.today())

if uploaded_file is not None:
    st.success(f"Uploaded file: {uploaded_file.name}")

if st.button("Generate Documents", type="primary"):
    if uploaded_file is None:
        st.error("Please upload an Excel file first.")
    elif not invoice_no.strip():
        st.error("Please enter an invoice number.")
    elif not course_title.strip():
        st.error("Please enter a course title.")
    else:
        try:
            output_dir = Path("generated_output")
            service = GenerationService(base_output_dir=output_dir)
            generated_files = service.generate_from_file(
                file_obj=uploaded_file,
                invoice_no=invoice_no.strip(),
                course_title=course_title.strip(),
                training_from=training_from,
                training_to=training_to,
                invoice_date=invoice_date,
                letter_date=letter_date,
            )
            zip_buffer = io.BytesIO()
            with ZipFile(zip_buffer, "w", ZIP_DEFLATED) as zip_file:
                for file_path in generated_files:
                    archive_name = file_path.relative_to(output_dir)
                    zip_file.write(file_path, arcname=str(archive_name))
            zip_buffer.seek(0)
            st.success(f"Generated {len(generated_files)} document(s) successfully.")
            st.download_button(label="Download ZIP", data=zip_buffer, file_name="cpr_generated_documents.zip", mime="application/zip")
        except Exception as e:
            st.error(f"Error: {str(e)}")
