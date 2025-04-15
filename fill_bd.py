# fill_bd.py
import sys
import os
from docxtpl import DocxTemplate
from datetime import datetime
from parse_rows import parse_bd_row
from airline_map import AIRLINE_MAP

def get_bd_template_path():
    return "BD_Sheet_Template.docx"

def get_output_folder():
    # same logic or import from main_gui
    # but let's define locally for convenience
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.getcwd()
    out_dir = os.path.join(exe_dir, "GeneratedDocuments")
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
    return out_dir

def generate_bd_sheet(record):
    current_dt = datetime.now().strftime("%m/%d/%Y %H:%M")
    template_path = get_bd_template_path()
    try:
        doc = DocxTemplate(template_path)
    except Exception as e:
        print(f"Error opening template '{template_path}': {e}")
        return None

    carriers = record.get("carriers", [])
    c1_name = carriers[0][0] if len(carriers)>0 else ""
    c1_ctn  = carriers[0][1] if len(carriers)>0 else ""
    c2_name = carriers[1][0] if len(carriers)>1 else ""
    c2_ctn  = carriers[1][1] if len(carriers)>1 else ""
    c3_name = carriers[2][0] if len(carriers)>2 else ""
    c3_ctn  = carriers[2][1] if len(carriers)>2 else ""

    context = {
        "CURRENT_DATETIME": current_dt,
        "MAWBs": record.get("mawb",""),
        "PCS": record.get("pieces",""),
        "PMC": record.get("pmc",""),
        "LAST_MILE": record.get("last_mile",""),
        "HOLDNUMBERS": record.get("hold",""),
        "carrier1": c1_name,
        "ctn1": c1_ctn,
        "carrier2": c2_name,
        "ctn2": c2_ctn,
        "carrier3": c3_name,
        "ctn3": c3_ctn,
    }
    doc.render(context)

    out_folder = get_output_folder()
    out_filename = os.path.join(out_folder, f"BD_{record.get('mawb','')}.docx")
    doc.save(out_filename)
    print(f"Generated: {out_filename}")
    return out_filename

def main():
    raw_text = input("Paste exactly one row for the B/D sheet:\n")
    # parse row, call generate_bd_sheet, etc.
    # ...
    pass
