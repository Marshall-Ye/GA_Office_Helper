# fill_ptt.py
import sys
import os
from docxtpl import DocxTemplate
from datetime import date
from airline_map import AIRLINE_MAP
from parse_rows import parse_ptt_rows_from_text
import docx2pdf
from docx2pdf import convert


def get_template_path():
    return "BD_Sheet_Template.docx"
...


def get_template_path():
    return "LAX_PTT_Template.docx"


def generate_ptt_for_records(records):
    """
    For a list of records (each containing keys: 'mawb', 'flt', 'pieces', 'weight'),
    fill in your PTT template for each record.

    Returns a list of output filenames.
    """
    today_str = date.today().strftime("%m/%d/%Y")
    output_files = []

    for record in records:
        mawb = record["mawb"]
        flt = record["flt"]
        pieces = record["pieces"]
        weight = record["weight"]
        prefix = mawb[:3]
        airline_name = AIRLINE_MAP.get(prefix, "Unknown Airline")

        doc = DocxTemplate(get_template_path())
        context = {
            "MAWB": mawb,
            "PIECES": pieces,
            "WEIGHT": weight,
            "FLT": flt,
            "AirlineName": airline_name,
            "TODAYS_DATE": today_str
        }
        doc.render(context)

        out_filename = os.path.abspath(f"PTT_{mawb}.docx")
        doc.save(out_filename)
        print(f"Generated: {out_filename}")
        output_files.append(out_filename)
    return output_files


# Optional: A main() for console usage if desired.
if __name__ == "__main__":
    # You could support console reading here:
    raw = input("Paste your rows here (end with an empty line):\n")
    # ...
    records = parse_ptt_rows_from_text(raw)
    generate_ptt_for_records(records)
