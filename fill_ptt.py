"""
fill_ptt.py – backend helpers for GA Office Helper
==================================================

• Locates the PTT Word template inside the bundled `_internal/` folder
  (works both from source and from a PyInstaller build).
• Defines firm metadata (GA3 / Pinto) and exposes
  `generate_ptt_for_records()` which turns parsed rows into PDF files.
"""

from __future__ import annotations

import os
import sys
import subprocess
from datetime import date
from pathlib import Path
from typing import List

from docxtpl import DocxTemplate
from docx2pdf import convert
from airline_map import AIRLINE_MAP          # local module

# ───────────────────── Windows-COM helper ────────────────────────────
if sys.platform == "win32":
    import pythoncom

    def _com_begin() -> None:
        """Must be called once **inside** every worker-thread that uses Word."""
        pythoncom.CoInitialize()                 # balanced by CoUninitialize

    def _com_end() -> None:
        pythoncom.CoUninitialize()

else:  # macOS / Linux – docx2pdf won’t run anyway, but keep no-op stubs
    def _com_begin() -> None: ...
    def _com_end() -> None:   ...
# ─────────────────────────────────────────────────────────────────────

# ─────────────── firm lookup table ───────────────────────────────────
FIRM_MAP: dict[str, dict[str, str]] = {
    "GA3": {
        "FirmCode":      "WBS6",
        "FirmName":      "GOLDEN ARCUS",
        "FullFirmName":  "GOLDEN ARCUS INTERNATIONAL CORP",
        "Address":       "5343 W. Imperial Hwy Ste 700, Los Angeles CA 90045",
    },
    "Pinto": {
        "FirmCode":      "W353",
        "FirmName":      "PINTO EXPRESS",
        "FullFirmName":  "Pinto Express",
        "Address":       "1001 W Walnut St, Compton CA 90220",
    },
}
FIRM_CHOICES: tuple[str, ...] = tuple(FIRM_MAP.keys())   # ('GA3', 'Pinto')
DEFAULT_FIRM: str = "GA3"

# ─────────────── locate bundled template ─────────────────────────────
def get_template_path() -> str:
    if hasattr(sys, "_MEIPASS"):             # one-file exe
        base = Path(sys._MEIPASS)
    elif getattr(sys, "frozen", False):      # one-folder
        base = Path(sys.executable).parent
    else:                                    # source run
        base = Path(__file__).parent
    return str(base / "_internal" / "LAX_PTT_Template.docx")

# ─────────────── output-folder helpers ───────────────────────────────
def get_output_folder() -> str:
    base = Path(sys.executable).parent if getattr(sys, "frozen", False) \
           else Path.cwd()
    out = base / "GeneratedDocuments"
    out.mkdir(exist_ok=True)
    return str(out)

def open_output_folder() -> None:
    folder = get_output_folder()
    if sys.platform == "win32":
        os.startfile(folder)
    elif sys.platform == "darwin":
        subprocess.run(["open", folder], check=False)
    else:
        subprocess.run(["xdg-open", folder], check=False)

# ─────────────── main generator ──────────────────────────────────────
def generate_ptt_for_records(
    records: List[dict],
    firm_key: str = DEFAULT_FIRM,
    op_name: str = "",
) -> List[str]:
    """
    Convert parsed *records* into PDF PTTs and return the PDF paths.
    """
    firm   = FIRM_MAP.get(firm_key, FIRM_MAP[DEFAULT_FIRM])
    today  = date.today().strftime("%m/%d/%Y")
    outdir = get_output_folder()
    pdfs   = []

    _com_begin()                   # ← init COM for **this** thread

    try:
        for rec in records:
            mawb, flt = rec["mawb"], rec["flt"]
            pcs,  wt  = rec["pieces"], rec["weight"]
            airline   = AIRLINE_MAP.get(mawb[:3], "Unknown Airline")

            doc = DocxTemplate(get_template_path())
            doc.render({
                "MAWB": mawb, "PIECES": pcs, "WEIGHT": wt, "FLT": flt,
                "AirlineName": airline, "TODAYS_DATE": today,
                "FirmCode": firm["FirmCode"], "FirmName": firm["FirmName"],
                "FullFirmName": firm["FullFirmName"], "Address": firm["Address"],
                "OPName": op_name,
            })

            safe_mawb = "".join(ch for ch in mawb.splitlines()[0]
                                if ch not in r'\/:*?"<>|')
            docx_path = Path(outdir) / f"PTT_{safe_mawb}.docx"
            doc.save(docx_path)

            # Word → PDF
            try:
                convert(str(docx_path))
                pdf_path = docx_path.with_suffix(".pdf")
                pdfs.append(str(pdf_path))
                docx_path.unlink(missing_ok=True)   # keep PDFs only
            except Exception as e:                  # noqa: BLE001
                print(f"[WARN] PDF conversion failed for {mawb}: {e}")

    finally:
        _com_end()               # tidy up COM even if something exploded

    return pdfs

# ─────────────── CLI demo (optional) ─────────────────────────────────
if __name__ == "__main__":                         # pragma: no cover
    from parse_rows import parse_ptt_rows_from_text

    print("Paste PTT rows (blank line ends):")
    buf: list[str] = []
    while True:
        ln = input()
        if not ln.strip():
            break
        buf.append(ln)

    rows = parse_ptt_rows_from_text("\n".join(buf))
    op   = input("Operator name: ").strip()
    firm = input(f"Firm ({'/'.join(FIRM_CHOICES)}): ").strip() or DEFAULT_FIRM

    for p in generate_ptt_for_records(rows, firm, op):
        print("Saved", p)
