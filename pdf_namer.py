import fitz  # PyMuPDF
import re
import os
from datetime import datetime

def extract_info_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()

    # Extract Client ID
    client_id_match = re.search(r'File Number: ([A-Z\-0-9]+)', text)
    client_id = client_id_match.group(1) if client_id_match else "UnknownID"

    # Extract Client Names
    client_names_match = re.search(r'Homeowner Name\(s\): (.+)', text)
    client_names = client_names_match.group(1) if client_names_match else "UnknownClients"
    client_names_clean = client_names.replace(' ', '').replace('and', '_')

    # Determine Document Type
    if "Forward Sale Option and Exchange Agreement" in text:
        doc_type = "ForwardSaleAgreement"
    elif "Compliance Agreement" in text:
        doc_type = "ComplianceAgreement"
    elif "Affidavit" in text:
        doc_type = "Affidavit"
    else:
        doc_type = "GeneralDoc"

    # Extract Date
    date_match = re.search(r'Date: ([A-Za-z]+\s\d{1,2},\s\d{4})', text)
    if date_match:
        date_str = date_match.group(1)
        date = datetime.strptime(date_str, "%B %d, %Y").strftime("%m%d%Y")
    else:
        date = "UnknownDate"

    filename = f"{client_id}_{client_names_clean}_{doc_type}_{date}.pdf"
    return filename

def rename_and_move_pdf(original_path, destination_folder):
    new_filename = extract_info_from_pdf(original_path)
    new_path = os.path.join(destination_folder, new_filename)

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    os.rename(original_path, new_path)
    print(f"Renamed and moved to: {new_path}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python pdf_namer.py <path_to_pdf> <destination_folder>")
        sys.exit(1)

    pdf_file = sys.argv[1]
    destination_folder = sys.argv[2]

    rename_and_move_pdf(pdf_file, destination_folder)
