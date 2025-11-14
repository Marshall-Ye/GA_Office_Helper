"""
parse_rows.py
=============

Parse tab-delimited rows copied from Google Sheets / Excel that may contain

* **Quoted cells** with embedded new-lines (`Alt + Enter` in Excel).
* **Missing trailing columns** (we only rely on a subset anyway).

The module exports `parse_ptt_rows_from_text` which returns plain Python data
that the rest of GA Office Helper consumes.
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


