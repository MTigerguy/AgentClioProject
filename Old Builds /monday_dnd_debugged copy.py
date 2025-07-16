#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter, legal
from reportlab.pdfgen import canvas
from io import BytesIO
from datetime import datetime

# Debug startup
try:
    with open(os.path.expanduser("~/Desktop/agent_monday_launch.log"), "a") as log:
        log.write("Launching Agent Monday...\n")
except Exception as e:
    print(f"Debug log failed: {e}")

def create_blank_page(size):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=size)
    c.showPage()
    c.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0]

def classify_and_split_pdf(input_path, log_callback):
    reader = PdfReader(input_path)
    letter_writer = PdfWriter()
    legal_writer = PdfWriter()
    letter_with_blanks = PdfWriter()
    legal_with_blanks = PdfWriter()
    summary_lines = []

    letter_count = 0
    legal_count = 0

    summary_lines.append(f"Total pages: {len(reader.pages)}\n")
    log_callback(f"Total pages: {len(reader.pages)}")

    for i, page in enumerate(reader.pages):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)

        if abs(height - 792) < 5:
            letter_writer.add_page(page)
            legal_with_blanks.add_page(create_blank_page(legal))
            letter_with_blanks.add_page(page)
            legal_writer.add_page(create_blank_page(legal))
            letter_count += 1
            page_type = "Letter"
        elif abs(height - 1008) < 5:
            legal_writer.add_page(page)
            letter_with_blanks.add_page(create_blank_page(letter))
            legal_with_blanks.add_page(page)
            letter_writer.add_page(create_blank_page(letter))
            legal_count += 1
            page_type = "Legal"
        else:
            msg = f"Page {i+1}: Unknown size ({width}x{height})"
            summary_lines.append(msg + "\n")
            log_callback(msg)
            continue

        msg = f"Page {i+1}: {page_type}"
        summary_lines.append(msg + "\n")
        log_callback(msg)

    summary_lines.insert(1, f"Letter pages: {letter_count}\n")
    summary_lines.insert(2, f"Legal pages: {legal_count}\n")

    output_folder = os.path.join(os.path.expanduser("~/Documents"), "Mondays Files")
    os.makedirs(output_folder, exist_ok=True)

    original_name = os.path.splitext(os.path.basename(input_path))[0]
    today_str = datetime.now().strftime("%Y-%m-%d")

    def save(writer, name):
        with open(os.path.join(output_folder, f"{name}_{original_name}_{today_str}.pdf"), "wb") as f:
            writer.write(f)

    save(letter_writer, "LETTER")
    save(legal_writer, "LEGAL")
    save(letter_with_blanks, "LETTER_IN_ORDER")
    save(legal_with_blanks, "LEGAL_IN_ORDER")

    with open(os.path.join(output_folder, f"SUMMARY_{original_name}_{today_str}.log"), "w") as f:
        f.writelines(summary_lines)

    log_callback("Done! Files saved to Documents/Mondays Files.")
    os.system(f'open "{output_folder}"')

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        log_text.delete(1.0, tk.END)
        classify_and_split_pdf(file_path, lambda msg: log_text.insert(tk.END, msg + "\n"))

def drop_handler(event):
    file_path = event.data.strip().replace("{", "").replace("}", "")
    if file_path.lower().endswith(".pdf"):
        log_text.delete(1.0, tk.END)
        classify_and_split_pdf(file_path, lambda msg: log_text.insert(tk.END, msg + "\n"))
    else:
        messagebox.showerror("Invalid File", "Please drop a valid PDF file.")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title("Agent Monday – PDF Sorter")
    root.geometry("600x400")

    frame = tk.Frame(root)
    frame.pack(pady=20)

    btn = tk.Button(frame, text="Select PDF File", command=browse_file, font=("Arial", 14))
    btn.pack()

    log_label = tk.Label(root, text="Log Output:")
    log_label.pack()

    log_text = tk.Text(root, height=15, width=70)
    log_text.pack()

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', drop_handler)

    root.mainloop()
