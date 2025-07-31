import PySimpleGUI as sg
import fitz  # PyMuPDF
import pandas as pd
import re
import os

# Extract text from PDF
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

# Simplified extraction logic
def extract_om_summary(text):
    summary = {
        "Equipment Name / Type": "APGN Electric Blower" if "blower" in text.lower() else "Not found",
        "Manufacturer": "APG-Neuros" if "APG-Neuros" in text else "Not found",
        "Model Number": re.findall(r"P\d{2}-\d\.\d+MW-\d+", text),
        "Serial Number": re.findall(r"Serial No[:\s]+\S+", text),
        "Startup Procedure": "Check system power-up, PLC check" if "startup" in text.lower() else "Not found",
        "Shutdown Procedure": "Controlled stop via PLC/HMI" if "shutdown" in text.lower() else "Not found",
        "Preventive Maintenance Schedule": "Check for daily/weekly tasks" if "maintenance" in text.lower() else "Not found",
        "Troubleshooting Guide": "Search for alarm/fault list" if "troubleshooting" in text.lower() else "Not found",
    }
    return summary

# Save summary to Excel
def save_summary_to_excel(summary, output_path):
    df = pd.DataFrame(list(summary.items()), columns=["Data Point", "Extracted Content"])
    df.to_excel(output_path, index=False)

# GUI Layout
layout = [
    [sg.Text("Select O&M Manual (PDF):"), sg.Input(), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
    [sg.Text("Save Excel Summary As:"), sg.Input(default_text="summary_output.xlsx"), sg.FileSaveAs(file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Button("Run Extraction"), sg.Exit()]
]

window = sg.Window("O&M Manual Extractor", layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break
    if event == "Run Extraction":
        pdf_path = values[0]
        excel_path = values[1]
        if not os.path.exists(pdf_path):
            sg.popup("Error", "PDF file not found.")
            continue
        try:
            text = extract_text_from_pdf(pdf_path)
            summary = extract_om_summary(text)
            save_summary_to_excel(summary, excel_path)
            sg.popup("Success", f"Summary saved to {excel_path}")
        except Exception as e:
            sg.popup("Error", str(e))

window.close()
