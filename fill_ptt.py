import sys
import os
from docxtpl import DocxTemplate
from datetime import date
from parse_rows import parse_rows_from_input
from airline_map import AIRLINE_MAP

def get_template_path():
    # When running under PyInstaller, sys._MEIPASS is the temp folder
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, "LAX_PTT_Template.docx")
    else:
        # Fallback if you're running in normal Python (not in a .exe)
        return "LAX_PTT_Template.docx"

def main():

    today_str = date.today().strftime("%m/%d/%Y")

    # 1) Get a list of records from user input (via parse_rows.py)
    records = parse_rows_from_input()

    # 2) For each record, fill the Word template and save an output .docx
    for record in records:
        mawb = record["mawb"]
        flt = record["flt"]
        pieces = record["pieces"]
        weight = record["weight"]

        prefix = mawb[:3]
        airline_name = AIRLINE_MAP.get(prefix, "Unknown Airline")

        # Load your template
        doc = DocxTemplate(get_template_path())

        # Build the context for the placeholders in your template
        context = {
            "MAWB": mawb,
            "PIECES": pieces,
            "WEIGHT": weight,
            "FLT": flt,
            "AirlineName": airline_name,  # use your airline_name variable
            "TODAYS_DATE": today_str  # add your current date string
        }

        # Render and save
        doc.render(context)

        out_filename = f"PTT_{mawb}.docx"
        doc.save(out_filename)
        print(f"Generated: {out_filename}")


if __name__ == "__main__":
    main()