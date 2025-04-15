import tkinter as tk
from tkinter import ttk, messagebox, PhotoImage
from datetime import date
import os
import sys
from docx2pdf import convert

# parse_rows for PTT/BD data
from parse_rows import parse_ptt_rows_from_text, parse_bd_row
# fill_ptt for retrieving your PTT template path
from fill_ptt import get_template_path
# fill_bd for generating BD docs
from fill_bd import generate_bd_sheet

def on_closing():
    root.destroy()
    sys.exit(0)

def get_exe_folder():
    """
    Returns the folder where the .exe is located if frozen,
    otherwise returns the current working directory when running in an IDE.
    """
    if getattr(sys, 'frozen', False):  # if running from .exe
        return os.path.dirname(sys.executable)
    else:
        return os.getcwd()

def get_output_folder():
    """
    Returns the path to 'GeneratedDocuments' subfolder in the exe's folder.
    Creates it if it doesn't exist.
    """
    base_dir = get_exe_folder()
    out_dir = os.path.join(base_dir, "GeneratedDocuments")
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
    return out_dir

# -----------------------------
# Build main window
# -----------------------------
root = tk.Tk()
root.title("GA Office Helper")
root.protocol("WM_DELETE_WINDOW", on_closing)

try:
    # Just assume GA_Logo.png is next to the exe or in same folder if in IDE
    logo_img = PhotoImage(file="GA_Logo.png")
    logo_label = tk.Label(root, image=logo_img)
    logo_label.pack(pady=5)
except Exception as e:
    print(f"Could not load logo: {e}")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# -----------------------------
# TAB 1: PTT Generator
# -----------------------------
ptt_frame = ttk.Frame(notebook)
notebook.add(ptt_frame, text="PTT Generator")

ptt_instructions = ttk.Label(ptt_frame, text=(
    "Paste multiple PTT rows here.\n"
    "Then click 'Generate PTT Docs'."
))
ptt_instructions.pack(padx=10, pady=10)

ptt_text_area = tk.Text(ptt_frame, width=100, height=15)
ptt_text_area.pack(padx=10, pady=5)

ptt_progress_label = ttk.Label(ptt_frame, text="")
ptt_progress_label.pack(pady=2)

ptt_progress_bar = ttk.Progressbar(ptt_frame, orient="horizontal", length=300, mode="determinate")
ptt_progress_bar.pack(pady=2)

def generate_ptt_docs():
    """
    Reads multiple PTT rows, saves .docx to a 'GeneratedDocuments' subfolder,
    converts each to PDF, removes the docx, etc.
    """
    raw_text = ptt_text_area.get("1.0", tk.END)
    if not raw_text.strip():
        messagebox.showwarning("No Data", "Please paste some rows first.")
        return

    records = parse_ptt_rows_from_text(raw_text)
    record_count = len(records)
    if record_count == 0:
        messagebox.showwarning("No Data", "No valid PTT rows found.")
        return

    ptt_progress_bar["maximum"] = record_count
    ptt_progress_bar["value"] = 0

    generated_count = 0
    today_str = date.today().strftime("%m/%d/%Y")
    output_folder = get_output_folder()  # e.g. /path/to/exe/GeneratedDocuments

    for i, record in enumerate(records, start=1):
        mawb = record["mawb"]
        flt = record["flt"]
        pieces = record["pieces"]
        weight = record["weight"]

        from airline_map import AIRLINE_MAP
        prefix = mawb[:3]
        airline_name = AIRLINE_MAP.get(prefix, "Unknown Airline")

        from docxtpl import DocxTemplate
        template_path = get_template_path()  # "LAX_PTT_Template.docx"
        doc = DocxTemplate(template_path)

        context = {
            "MAWB": mawb,
            "PIECES": pieces,
            "WEIGHT": weight,
            "FLT": flt,
            "AirlineName": airline_name,
            "TODAYS_DATE": today_str
        }
        doc.render(context)

        # Save docx into our output folder
        out_docx = os.path.join(output_folder, f"PTT_{mawb}.docx")
        doc.save(out_docx)

        try:
            convert(out_docx)  # doc -> PDF in the same folder
        except Exception as e:
            messagebox.showerror("Conversion Error", f"Error converting {out_docx}:\n{e}")
            return

        # Optionally remove the docx if you only want the PDF
        if os.path.exists(out_docx):
            os.remove(out_docx)

        generated_count += 1
        ptt_progress_bar["value"] = i
        ptt_progress_label.config(text=f"Processing {i}/{record_count}…")
        root.update_idletasks()

    ptt_text_area.delete("1.0", tk.END)
    ptt_progress_label.config(text="")
    ptt_progress_bar["value"] = 0

    messagebox.showinfo("Done", f"Generated {generated_count} PDF(s) in '{output_folder}'")

ptt_button = ttk.Button(ptt_frame, text="Generate PTT Docs", command=generate_ptt_docs)
ptt_button.pack(pady=10)

# -----------------------------
# TAB 2: B/D Sheet Generator
# -----------------------------
bd_frame = ttk.Frame(notebook)
notebook.add(bd_frame, text="B/D Sheet Generator")

bd_instructions = ttk.Label(bd_frame, text=(
    "Paste exactly one row for B/D sheet.\n"
    "Then click 'Generate BD Doc'.\n"
    "Saved in 'GeneratedDocuments' subfolder."
))
bd_instructions.pack(padx=10, pady=10)

bd_text_area = tk.Text(bd_frame, width=100, height=5)
bd_text_area.pack(padx=10, pady=5)

def generate_bd_doc():
    """
    Single-row B/D approach, saving the docx in 'GeneratedDocuments' subfolder.
    """
    raw_text = bd_text_area.get("1.0", tk.END).strip()
    if not raw_text:
        messagebox.showwarning("No Data", "Please paste exactly one row for B/D sheet.")
        return

    record = parse_bd_row(raw_text)
    if not record.get("mawb"):
        messagebox.showwarning("No Data", "No valid BD row or missing MAWB.")
        return

    out_docx = generate_bd_sheet(record)  # We'll modify fill_bd to also place it in that subfolder
    if not out_docx or not os.path.exists(out_docx):
        messagebox.showinfo("Done", "No BD doc generated or file not found.")
        return

    hold_val = record["hold"]
    mawb_val = record["mawb"]
    carriers = record.get("carriers", [])

    is_usps_override = (
        len(carriers) == 3
        and carriers[0] == ("800-", "")
        and carriers[1] == ("807-", "")
        and carriers[2] == ("808-", "")
    )

    popup_lines = [
        f"Clearance: {hold_val}",
        f"MAWB: {mawb_val}"
    ]
    if is_usps_override:
        popup_lines.append("Please fill in details for USPS-CO")

    bd_text_area.delete("1.0", tk.END)
    messagebox.showinfo("B/D Done", "\n".join(popup_lines))

bd_button = ttk.Button(bd_frame, text="Generate BD Doc", command=generate_bd_doc)
bd_button.pack(pady=10)

root.mainloop()
