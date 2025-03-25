import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
from docxtpl import DocxTemplate
from airline_map import AIRLINE_MAP
from fill_ptt import get_template_path
import sys
from docx2pdf import convert
import os


def parse_rows_from_text(raw_text):
    """
    Similar to parse_rows_from_input(), but parses from a string.
    Splits on newlines, then splits each line by tab.
    Returns a list of record dicts: [ {mawb, flt, pieces, weight}, ... ]
    """
    lines = raw_text.strip().splitlines()
    parsed_records = []

    for row in lines:
        cols = row.split("\t")

        def safe_get(lst, idx):
            return lst[idx] if idx < len(lst) else ""

        # F=5, J=9, L=11, N=13
        mawb   = safe_get(cols, 5)
        flt    = safe_get(cols, 9)
        pieces = safe_get(cols, 11)
        weight = safe_get(cols, 13)

        if not mawb.strip():
            continue

        parsed_records.append({
            "mawb": mawb,
            "flt": flt,
            "pieces": pieces,
            "weight": weight
        })

    return parsed_records


def on_closing():
    """
    Called when user clicks the 'X' button.
    Ensures the Python process really ends.
    """
    root.destroy()   # close the window
    sys.exit(0)      # forcibly end the interpreter

def generate_ptt_docs():
    """
    Reads text from the GUI text widget, parses each row,
    then fills the Word template for each record.
    """
    raw_text = text_area.get("1.0", tk.END)
    if not raw_text.strip():
        messagebox.showwarning("No Data", "Please paste some rows first.")
        return

    # Parse
    records = parse_rows_from_text(raw_text)

    today_str = date.today().strftime("%m/%d/%Y")

    # Suppose we got the records
    record_count = len(records)
    if record_count == 0:
        messagebox.showwarning("No Data", "No valid rows found.")
        return

    # Configure the Progressbar
    progress_bar["maximum"] = record_count
    progress_bar["value"] = 0

    generated_count = 0

    for i, record in enumerate(records, start=1):
        mawb = record["mawb"]
        flt = record["flt"]
        pieces = record["pieces"]
        weight = record["weight"]

        prefix = mawb[:3]
        airline_name = AIRLINE_MAP.get(prefix, "Unknown Airline")

        # Fill the template
        try:
            doc = DocxTemplate(get_template_path())
        except FileNotFoundError:
            messagebox.showerror(
                "Template Not Found",
                "Cannot find 'LAX_PTT_Template.docx' in the current folder."
            )
            return

        context = {
            "MAWB": mawb,
            "PIECES": pieces,
            "WEIGHT": weight,
            "FLT": flt,
            "AirlineName": airline_name,
            "TODAYS_DATE": today_str
        }

        # 2) Save the docx
        doc.render(context)
        out_docx = f"PTT_{mawb}.docx"
        doc.save(out_docx)

        # 3) Convert to PDF using docx2pdf
        convert(out_docx)  # This creates "PTT_{mawb}.pdf"

        # 4) (Optional) Delete the .docx if you only want a PDF
        os.remove(out_docx)

        generated_count += 1

        # Update progress
        progress_bar["value"] = i
        progress_label.config(text=f"Processing {i}/{record_count}…")
        root.update_idletasks()  # or root.update()

    # After the loop finishes:
    progress_label.config(text="")
    progress_bar["value"] = 0
    messagebox.showinfo("Done", f"Generated {generated_count} document(s)!")


# --- Build the simple Tkinter GUI ---
root = tk.Tk()
root.title("PTT Auto-Generator")

# Tie the window’s close (X button) to on_closing()
root.protocol("WM_DELETE_WINDOW", on_closing)

progress_label = ttk.Label(root, text="")
progress_label.pack(pady=2)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=2)


instructions = ttk.Label(root, text=(
    "Paste your spreadsheet rows here. "
    "Make sure Pick up Location is filled, then click 'Generate PTT Docs'."
))
instructions.pack(padx=10, pady=10)

text_area = tk.Text(root, width=100, height=20)
text_area.pack(padx=10, pady=5)

btn_generate = ttk.Button(root, text="Generate PTT Docs", command=generate_ptt_docs)
btn_generate.pack(pady=10)

root.mainloop()