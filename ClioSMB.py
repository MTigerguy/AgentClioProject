import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, simpledialog
from PyPDF2 import PdfReader, PdfWriter
import fitz      # type: ignore # PyMuPDF
import re
import datetime
import traceback
from openpyxl import Workbook, load_workbook # type: ignore
import csv

# --- Config handling for "dad version" ---
CONFIG_PATH = os.path.expanduser("~/.clio_config.json")

def get_base_dir():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                cfg = json.load(f)
            if "base_dir" in cfg and os.path.exists(cfg["base_dir"]):
                return cfg["base_dir"]
        except Exception:
            pass  # Ignore config errors and reprompt

    # Prompt for folder on first run
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Where do you want to save your Clio files?")
    root.destroy()
    if not folder:
        messagebox.showerror("No folder selected", "No folder chosen, app will exit.")
        raise SystemExit
    # Save to config
    with open(CONFIG_PATH, "w") as f:
        json.dump({"base_dir": folder}, f)
    return folder

BASE_DIR = get_base_dir()
LOG_EXCEL = os.path.join(BASE_DIR, "Clio_Log.xlsx")
LOG_CSV = os.path.join(BASE_DIR, "Clio_Log.csv")
ERROR_LOG = os.path.join(BASE_DIR, "Clio_app_error.log")

# Summary TXT creation
def create_summary_txt(folder, base, borrower, date_of_signing, letter_count, legal_count, other_count, total_count):
    filename = os.path.join(folder, f"{base}_Summary.txt")
    with open(filename, "w", encoding="utf-8") as f:
        f.write("Document Summary\n")
        f.write(f"Date Created: {datetime.date.today().strftime('%Y-%m-%d')}\n")
        f.write(f"Date of Signing: {date_of_signing}\n")
        f.write(f"Borrower(s): {borrower}\n")
        f.write(f"Total Page Count: {total_count}\n")
        f.write(f"  • Letter: {letter_count}\n")
        f.write(f"  • Legal: {legal_count}\n")
        f.write(f"  • Other: {other_count}\n")
    print(f"  Saving summary: {filename}")

#show status window
def show_status_window(text, filenames=None):
    win = tk.Toplevel()
    win.title("Processing Results")
    win.geometry("500x400")
    st = scrolledtext.ScrolledText(win, wrap=tk.WORD)
    st.pack(fill=tk.BOTH, expand=True)
    st.insert(tk.END, text)
    # Add output filenames, if provided
    if filenames:
        st.insert(tk.END, "\n\nOutput Files Created:\n")
        for fn in filenames:
            st.insert(tk.END, f"  {fn}\n")
    st.config(state=tk.DISABLED)

def select_folder_and_name():
    parent = tk._default_root if tk._default_root else tk.Toplevel()
    folder = filedialog.askdirectory(title="Select Output Folder", parent=parent)
    if not folder:
        if not tk._default_root:
            parent.destroy()
        return None, None
    def on_ok():
        name = name_var.get().strip()
        if not name or not name.replace("_", "").replace("-", "").isalnum() or len(name) > 20:
            messagebox.showerror("Invalid Name", "Enter 1-20 letters/numbers (_ and - allowed).", parent=dialog)
            return
        dialog.result = name
        dialog.destroy()
    dialog = tk.Toplevel(parent)
    dialog.title("Enter Base Filename")
    tk.Label(dialog, text="Base filename (1–20 chars, letters/numbers/_/-):").pack(padx=10, pady=8)
    name_var = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=name_var, width=25)
    entry.pack(padx=10, pady=4)
    entry.focus()
    tk.Button(dialog, text="OK", command=on_ok).pack(pady=10)
    dialog.result = None
    dialog.transient(parent)
    dialog.grab_set()
    parent.wait_window(dialog)
    if not tk._default_root:
        parent.destroy()
    if dialog.result:
        return folder, dialog.result
    return None, None

def split_and_save_pdfs(path, folder, base, reader):
    letter_writer = PdfWriter()
    legal_writer = PdfWriter()
    other_writer = PdfWriter()

    letter_count = legal_count = other_count = 0
    in_legal_block = False

    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        tp = get_paper_type(w, h)

        if tp == "Letter":
            letter_writer.add_page(page)
            letter_count += 1
            in_legal_block = False
        elif tp == "Legal":
            legal_writer.add_page(page)
            legal_count += 1
            if not in_legal_block:
                letter_writer.add_blank_page(width=8.5*72, height=11*72)
                in_legal_block = True
        else:
            other_writer.add_page(page)
            other_count += 1

    if letter_count or in_legal_block:
        print(f"  Saving: {os.path.join(folder, f'{base}_Letter.pdf')}")
        with open(os.path.join(folder, f"{base}_Letter.pdf"), "wb") as f:
            letter_writer.write(f)
    if legal_count:
        print(f"  Saving: {os.path.join(folder, f'{base}_Legal.pdf')}")
        with open(os.path.join(folder, f"{base}_Legal.pdf"), "wb") as f:
            legal_writer.write(f)
    if other_count:
        print(f"  Saving: {os.path.join(folder, f'{base}_Other.pdf')}")
        with open(os.path.join(folder, f"{base}_Other.pdf"), "wb") as f:
            other_writer.write(f)
    full_writer = PdfWriter()
    for page in reader.pages:
        full_writer.add_page(page)
    print(f"  Saving: {os.path.join(folder, f'{base}_Full.pdf')}")
    with open(os.path.join(folder, f"{base}_Full.pdf"), "wb") as f:
        full_writer.write(f)

    date_of_signing = simpledialog.askstring("Date of Signing", "Enter date of signing (YYYY-MM-DD):")
    borrower_name = base.split("_")[0]
    total_count = len(reader.pages)
    create_summary_txt(folder, base, borrower_name, date_of_signing, letter_count, legal_count, other_count, total_count)

    return f"{os.path.basename(path)} → Letter:{letter_count}, Legal:{legal_count}, Other:{other_count}\n"

def process_pdfs_individually_with_filelist(file_paths):
    if not file_paths:
        return

    log_txt = ""
    output_filenames = []
    for path in file_paths:
        name = os.path.basename(path)
        print(f"Processing: {path}")
        try:
            folder, base = extract_base_filename(path)
            print(f"  Extracted folder: {folder}")
            print(f"  Extracted base: {base}")
            if not folder or not base:
                skip_reason = f"{os.path.basename(path)} → SKIPPED (No name/folder)"
                print(f"  {skip_reason}")
                log_txt += skip_reason + "\n"
                continue
            os.makedirs(folder, exist_ok=True)
            reader = PdfReader(path)
            summary = split_and_save_pdfs(path, folder, base, reader)
            log_txt += summary
            for typ in ["Letter", "Legal", "Other", "Full"]:
                out_path = os.path.join(folder, f"{base}_{typ}.pdf")
                if os.path.exists(out_path):
                    output_filenames.append(out_path)
        except Exception as e:
            print(f"  ERROR processing {path}: {e}")
            log_txt += f"{os.path.basename(path)} → ERROR: {str(e)}\n"
            continue

    show_status_window(log_txt, filenames=output_filenames)
    log_action("Process", [os.path.basename(f) for f in file_paths], log_txt)
    print("Batch processing complete. See above for details.")

def show_intake_window():
    intake_win = tk.Tk()
    intake_win.title("Clio Document Intake")
    intake_win.geometry("900x350")

    selected_files = []

    def select_files():
        files = filedialog.askopenfilenames(
            title="Select PDF Files", filetypes=[("PDF files", "*.pdf")]
        )
        if files:
            selected_files.clear()
            selected_files.extend(files)
            file_listbox.delete(0, tk.END)
            for f in files:
                file_listbox.insert(tk.END, f)

    def process_files():
        if not selected_files:
            messagebox.showerror("No Files", "Please select PDFs first.")
            return
        process_pdfs_individually_with_filelist(selected_files)

    def refresh():
        selected_files.clear()
        file_listbox.delete(0, tk.END)

    def exit_app():
        intake_win.destroy()

    file_listbox = tk.Listbox(intake_win, width=100, height=10)
    file_listbox.pack(pady=10)

    tk.Button(intake_win, text="Select PDFs", command=select_files).pack()

    btn_frame = tk.Frame(intake_win)
    btn_frame.pack(pady=20)

    tk.Button(btn_frame, text="Process", width=15, command=process_files).grid(row=0, column=0, padx=5)
    tk.Button(btn_frame, text="Refresh", width=15, command=refresh).grid(row=0, column=1, padx=5)
    tk.Button(btn_frame, text="Exit", width=15, command=exit_app).grid(row=0, column=2, padx=5)

    intake_win.mainloop()

def manual_name_prompt():
    class NamePrompt(simpledialog.Dialog):
        def body(self, master):
            tk.Label(master, text="Last Name (max 10 letters):").grid(row=0)
            tk.Label(master, text="First Initial:").grid(row=1)
            self.last_var = tk.StringVar()
            self.init_var = tk.StringVar()
            self.last_entry = tk.Entry(master, textvariable=self.last_var)
            self.init_entry = tk.Entry(master, textvariable=self.init_var, width=4)
            self.last_entry.grid(row=0, column=1)
            self.init_entry.grid(row=1, column=1)
            return self.last_entry

        def validate(self):
            last = self.last_var.get().strip()
            initial = self.init_var.get().strip().upper()
            print(f"DEBUG: last='{last}', initial='{initial}'")
            if not last or len(last) > 10 or not last.isalpha():
                messagebox.showerror("Error", "Last name must be 1–10 letters.", parent=self)
                return False
            if not initial or not (1 <= len(initial) <= 2) or not initial.isalpha():
                messagebox.showerror("Error", "First initial must be 1 or 2 letters.", parent=self)
                return False
            return True

        def apply(self):
            self.result = (self.last_var.get().strip().capitalize(), self.init_var.get().strip().upper())

    parent = tk._default_root if tk._default_root else tk.Toplevel()
    prompt = NamePrompt(parent, "Filename: Cannot find name. Please provide last name (max 10 letters) and first initial.")
    if not tk._default_root:
        parent.destroy()
    if hasattr(prompt, "result") and prompt.result and prompt.result[0] and prompt.result[1]:
        return f"{prompt.result[0]}_{prompt.result[1]}"
    else:
        return None

def extract_base_filename(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""

    instruction_filters = [
        "attention closing agent",
        "closing agent instruction",
        "instructions to the closing agent",
        "attention signing agent",
        "notary checklist",
        "special instructions",
        "agent acknowledgement",
    ]

    not_borrower_filters = [
        "lender", "lenders", "lender name", "lender rep", "lender representative",
        "mortgage company", "mortgage companies", "bank", "banks", "servicer", "servicers",
        "broker", "brokers", "realty", "real estate", "real estate agent", "real estate agents",
        "title", "title agent", "title agents", "title rep", "title reps", "title company", "title companies",
        "escrow officer", "escrow officers", "settlement agent", "settlement agents", "notary", "notaries",
        "witness", "witnesses", "signature", "signatures", "loan officer", "loan officers", "processor", "processors",
        "signing agent", "signing agents", "closing agent", "closing agents", "underwriter", "underwriters",
        "attorney", "attorneys", "law firm", "law firms", "company", "companies", "inc", "inc.", "llc", "llc.", "corp",
        "corporation", "corporations", "co.", "co", "llp", "llp.", "pllc", "pllc.", "plc", "plc.", "pa", "p.a.", "pc", "p.c.",
        "associates", "associate", "group", "groups", "firm", "firms", "office", "offices", "department", "departments",
        "division", "divisions", "section", "sections", "admin", "administrator", "administrators", "administer",
        "president", "presidents", "vice president", "secretary", "secretaries", "manager", "managers", "management",
        "director", "directors", "officer", "officers", "official", "officials", "contact", "contacts", "employee", "employees",
        "staff", "team", "teams", "counsel", "adviser", "advisor", "consultant", "consultants", "independent contractor",
        "independent contractors", "contractor", "contractors", "organizer", "organizers", "participant", "participants",
        "benefactor", "benefactors", "grantor", "grantors", "grantee", "grantees", "remitter", "remitters",
        "payee", "payees", "payor", "payors", "mortgagor", "mortgagors", "mortgagee", "mortgagees",
        "trust", "trusts", "trustee", "trustees", "foundation", "foundations", "estate", "estates", "heir", "heirs",
        "beneficiary", "beneficiaries", "power of attorney", "poa", "personal representative", "personal representatives",
        "successor", "successors", "authorized", "representative", "representatives", "agent", "agents", "approved",
        "accept", "accepted", "seller", "sellers", "buyer", "buyers", "borrower", "borrowers", "co-borrower", "co-borrowers",
        "joint tenant", "joint tenants", "spouse", "spouses", "partner", "partners", "appointee", "appointees",
        "recipient", "recipients", "customer", "customers", "client", "clients", "occupant", "occupants",
        "landlord", "landlords", "tenant", "tenants", "lessee", "lessees", "lessor", "lessors", "guarantor", "guarantors",
        "north charleston", "charleston", "clearedge", "deceased", "account", "accounts", "section", "sections",
        "division", "divisions", "administer", "controller", "supervisor", "supervisors", "applicant", "applicants"
    ]
    generic_labels = {"borrower", "owner", "seller", "buyer", "applicant", "customer", "client"}

    for i in range(min(20, doc.page_count)):
        page_text = doc[i].get_text()
        page_start = page_text[:300].lower()
        if any(filter_text in page_start for filter_text in instruction_filters):
            if ("homeowner name" in page_text.lower()) or ("borrower" in page_text.lower()) or ("applicant" in page_text.lower()):
                text += page_text
                continue
            continue
        text += page_text

    def is_valid_name(candidate):
        candidate_lower = candidate.lower().strip()
        if candidate_lower in generic_labels:
            return False
        if any(nb in candidate_lower for nb in not_borrower_filters):
            return False
        if re.search(r"[^a-zA-Z ,.'&-]", candidate):
            return False
        if len(candidate) < 3:
            return False
        return True

    name = None
    patterns = [
        r"Borrower Information\s*[:\n]+([A-Za-z ,.'&-]+)",
        r"Borrower\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Borrower:\s*([A-Za-z ,.'&-]+)",
        r"Homeowner Name\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Owner\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Property Owner\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Seller\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Buyer\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Applicant\(s\):\s*([A-Za-z ,.'&-]+)",
        r"Client:\s*([A-Za-z ,.'&-]+)",
        r"Customer:\s*([A-Za-z ,.'&-]+)",
    ]

    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            candidate = m.group(1).split(' and ')[0].split(',')[0].strip()
            candidate_lower = candidate.lower().strip()
            if candidate_lower in generic_labels or candidate_lower in not_borrower_filters:
                continue
            if is_valid_name(candidate):
                name = candidate