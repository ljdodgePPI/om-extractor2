import fitz  # PyMuPDF
import re
import pandas as pd

# Load the PDF
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

# Sample keyword-based extraction (simplified for demonstration)
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

# Save to Excel
def save_summary_to_excel(summary, output_path):
    df = pd.DataFrame(list(summary.items()), columns=["Data Point", "Extracted Content"])
    df.to_excel(output_path, index=False)

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python om_extractor.py <input_pdf_path> <output_excel_path>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    excel_path = sys.argv[2]

    text = extract_text_from_pdf(pdf_path)
    summary = extract_om_summary(text)
    save_summary_to_excel(summary, excel_path)
    print(f"Summary saved to {excel_path}")
