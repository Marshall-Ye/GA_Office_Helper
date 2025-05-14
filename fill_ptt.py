# fill_ptt.py  –  backend helpers for GA Office Helper
# ----------------------------------------------------
import os, sys, subprocess
from datetime import date
from pathlib import Path
from typing import List

from docxtpl  import DocxTemplate
from docx2pdf import convert

from airline_map import AIRLINE_MAP


# ─────────────────── template path ────────────────────
def get_template_path() -> str:
    base = Path(sys.executable).parent if getattr(sys, 'frozen', False) \
           else Path(__file__).parent
    return str(base / "_internal" / "LAX_PTT_Template.docx")




# ─────────────────── output folder helpers ────────────
def get_output_folder() -> str:
    """Return <exe_dir>/GeneratedDocuments  (create if missing)."""
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path.cwd()
    out  = base / "GeneratedDocuments"
    out.mkdir(exist_ok=True)
    return str(out)


def open_output_folder() -> None:
    """Open the GeneratedDocuments folder in system file-explorer."""
    folder = get_output_folder()
    if sys.platform == "win32":
        os.startfile(folder)
    elif sys.platform == "darwin":
        subprocess.run(["open", folder], check=False)
    else:
        subprocess.run(["xdg-open", folder], check=False)


# ─────────────────── main generator ───────────────────
def generate_ptt_for_records(
    records: List[dict],
    firm_code: str = "WBS6",
    op_name: str = ""
) -> List[str]:
    """
    For each record in *records* generate a PTT PDF.
    `records` must contain keys: mawb, flt, pieces, weight.

    Returns list of generated PDF filepaths.
    """
    today   = date.today().strftime("%m/%d/%Y")
    out_dir = get_output_folder()
    pdf_paths = []

    for rec in records:
        mawb, flt, pcs, wt = rec["mawb"], rec["flt"], rec["pieces"], rec["weight"]
        airline            = AIRLINE_MAP.get(mawb[:3], "Unknown Airline")

        # 1) render docx
        doc = DocxTemplate(get_template_path())
        doc.render({
            "MAWB": mawb,
            "PIECES": pcs,
            "WEIGHT": wt,
            "FLT": flt,
            "AirlineName": airline,
            "TODAYS_DATE": today,
            "FirmCode": firm_code,
            "OPName": op_name
        })

        docx_path = Path(out_dir) / f"PTT_{mawb}.docx"
        doc.save(docx_path)

        # 2) convert to PDF
        try:
            convert(str(docx_path))                     # creates same-name PDF
            pdf_path = docx_path.with_suffix(".pdf")
            pdf_paths.append(str(pdf_path))
            docx_path.unlink(missing_ok=True)           # keep PDFs only
        except Exception as e:
            print(f"[WARN] PDF conversion failed for {mawb}: {e}")

    return pdf_paths


# ─────────────────── CLI test harness (optional) ───────
if __name__ == "__main__":
    from parse_rows import parse_ptt_rows_from_text

    raw = []
    print("Paste your PTT rows (end with blank line):")
    while True:
        line = input()
        if not line.strip():
            break
        raw.append(line)
    text = "\n".join(raw)

    recs = parse_ptt_rows_from_text(text)
    name = input("Operator name: ").strip()
    code = input("Firm code (WBS6/W353): ").strip() or "WBS6"

    pdfs = generate_ptt_for_records(recs, code, name)
    print("Generated:", *pdfs, sep="\n  ")
