"""
parse_rows.py
=============

Parse tab-delimited rows copied from Google Sheets / Excel that may contain

* **Quoted cells** with embedded new-lines (`Alt + Enter` in Excel).
* **Missing trailing columns** (we only rely on a subset anyway).

The module exports two helpers:

    parse_ptt_rows_from_text(raw_clipboard_text)  ->  list[dict]
    parse_bd_row(raw_clipboard_text)              ->  dict

Both return *plain* Python data that the rest of GA Office Helper consumes.
"""

from __future__ import annotations
import csv
from io import StringIO
from typing import List, Dict


# ---------------------------------------------------------------------------
def _csv_rows(raw: str) -> List[List[str]]:
    """
    Convert clipboard text to a list of rows (each row is a list of columns).

    * delimiter = TAB
    * quotechar = "
    * embedded newlines inside quotes are kept inside the cell
    """
    reader = csv.reader(
        StringIO(raw),
        delimiter="\t",
        quotechar='"',
        quoting=csv.QUOTE_MINIMAL,
        skipinitialspace=False,
        strict=False,
    )
    return list(reader)


def _safe(cols: List[str], idx: int) -> str:
    """Return cols[idx].strip() or an empty string if missing."""
    return cols[idx].strip() if idx < len(cols) else ""


# ---------------------------------------------------------------------------
#  PTT rows (MULTIPLE)
# ---------------------------------------------------------------------------
def parse_ptt_rows_from_text(raw_text: str) -> List[Dict[str, str]]:
    """
    Extract MAWB / Flight / Pieces / Weight from *every* valid row.

    Expected column indexes (0-based) in the clipboard export:

        5  MAWB
        9  Flight
       11  Pieces
       13  Weight
    """
    records: list[dict[str, str]] = []

    for cols in _csv_rows(raw_text):
        if len(cols) < 14:                        # critical columns missing
            continue

        mawb = _safe(cols, 5)
        if not mawb:                             # skip blank MAWB rows
            continue

        records.append({
            "mawb":   mawb,
            "flt":    _safe(cols, 9),
            "pieces": _safe(cols, 11),
            "weight": _safe(cols, 13),
        })

    return records


# ---------------------------------------------------------------------------
#  B/D sheet (SINGLE logical row)
# ---------------------------------------------------------------------------
def parse_bd_row(raw_text: str) -> Dict[str, str]:
    """
    Parse exactly ONE business row for the B/D generator.

    Column layout we care about (0-based):

         5  MAWB                  11  Pieces            15  PMC
        20  Hold                 21  Last-mile          22–25 Carriers (W–Z)
        26  USPS override flag   (a non-empty cell signals USPS-CO)

    * If MAWB column is empty → returns {}.
    * Embedded newlines inside PMC are replaced by ', ' for prettiness.
    """
    # Find the *first* row with a non-empty MAWB
    for cols in _csv_rows(raw_text):
        mawb = _safe(cols, 5)
        if not mawb:
            continue                                 # keep searching

        hold_str = _safe(cols, 20)
        if hold_str.isdigit():                       # "2" → "2 HOLD"
            hold_str += " HOLD"

        # --- basic fields ---------------------------------------------------
        record: dict[str, str] = {
            "mawb":       mawb,
            "pieces":     _safe(cols, 11),
            "pmc":        _safe(cols, 15).replace("\n", ", "),
            "hold":       hold_str,
            "last_mile":  _safe(cols, 21),
            "carriers":   [],
        }

        # --- carriers logic -------------------------------------------------
        if _safe(cols, 26):                          # USPS-CO override
            record["carriers"] = [("800-", ""), ("807-", ""), ("808-", "")]
        else:
            carrier_cols = {
                "YUN2": 22,   # W
                "UPS":  23,   # X
                "UNI":  24,   # Y
                "YWE":  25,   # Z
            }
            record["carriers"] = [
                (name, _safe(cols, idx))
                for name, idx in carrier_cols.items()
                if _safe(cols, idx)
            ][:3]                                     # max three entries

        return record

    # No usable row found
    return {}
