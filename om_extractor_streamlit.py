import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
import os

# --- Section Patterns ---
section_patterns = {
    "Equipment Name / Tag": r"(Equipment\s+Name|Tag)\s*[:\-]?\s*(.*)",
    "Manufacturer": r"(Manufacturer|Vendor|Supplier)\s*[:\-]?\s*(.*)",
    "Model / Serial Number": r"(Model\s*(Number)?|Serial\s*(Number)?)\s*[:\-]?\s*(.*)",
    "Installation Instructions": r"(Installation\s+Instructions|Installation\s+Procedure)",
    "Startup Procedure": r"(Startup\s+Procedure|Initial\s+Startup|Start[- ]?up\s+Instructions?)",
    "Normal Operation": r"(Normal\s+Operation|Operating\s+Instructions?|Operation\s+Procedure)",
    "Maintenance Schedule": r"(Maintenance\s+Schedule|Maintenance\s+Plan|Preventive\s+Maintenance)",
    "Troubleshooting Guide": r"(Troubleshooting\s+Guide|Problem\s+Resolution|Diagnostics?)",
    "Spare Parts List": r"(Spare\s+Parts\s+List|Recommended\s+Spare\s+Parts)",
    "Safety Warnings": r"(Safety\s+Warnings|Caution|Hazards|Safety\s+Instructions?)",
    "Electrical/Mechanical Specs": r"(Electrical\s+Specifications?|Mechanical\s+Specifications?)",
    "Drawings / Schematics": r"(Drawings|Schematics|Diagrams)"
}

# --- PDF Text Extractor ---
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

# --- Section Extractor ---
def extract_all_sections(pdf_path):
    text = extract_text_from_pdf(pdf_path)
    extracted = {"Section": [], "Content": []}

    for section, pattern in section_patterns.items():
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            content = "\n".join([m[-1] if isinstance(m, tuple) else m for m in matches])
            extracted["Section"].append(section)
            extracted["Content"].append(content)

    return extracted

# --- Streamlit Interface ---
st.title("📄 O&M Manual Extractor")

uploaded_file = st.file_uploader("Upload O&M Manual PDF", type="pdf")

if uploaded_file is not None:
    with open("temp_manual.pdf", "wb") as f:
        f.write(uploaded_file.read())

    st.info("Extracting... Please wait.")
    extracted_data = extract_all_sections("temp_manual.pdf")

    if extracted_data:
        df = pd.DataFrame(extracted_data)
        st.success("✅ Extraction complete! View and download below:")
        st.dataframe(df)

        output_file = "Blower_OM_summary_output.xlsx"
        df.to_excel(output_file, index=False)
        with open(output_file, "rb") as f:
            st.download_button("📥 Download Excel File", f, file_name=output_file)
    else:
        st.warning("⚠️ No content extracted.")