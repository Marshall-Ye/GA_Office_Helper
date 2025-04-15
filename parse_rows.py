# parse_rows.py

def safe_get(lst, idx):
    """
    Return the element at 'idx' if present, else an empty string.
    """
    return lst[idx] if (idx < len(lst)) else ""


# -----------------------------------------------------------------------
# PTT Parsing (Multiple Rows)
# -----------------------------------------------------------------------
def parse_ptt_rows_from_text(raw_text):
    """
    Parses multiple PTT rows (tab-separated) from a multi-line string.
    For each line, it extracts:
      - MAWB from column F (index 5)
      - Flight No. from column J (index 9)
      - Pieces from column L (index 11)
      - Weight from column N (index 13)

    Returns a list of dictionaries with keys:
      'mawb', 'flt', 'pieces', 'weight'

    This simpler approach assumes PTT rows do not contain embedded newlines.
    """
    lines = raw_text.strip().splitlines()
    records = []
    for line in lines:
        cols = line.split("\t")
        if len(cols) < 14:
            continue
        mawb = safe_get(cols, 5).strip()
        flt = safe_get(cols, 9).strip()
        pieces = safe_get(cols, 11).strip()
        weight = safe_get(cols, 13).strip()

        if not mawb:  # skip if no MAWB
            continue

        records.append({
            "mawb": mawb,
            "flt": flt,
            "pieces": pieces,
            "weight": weight
        })
    return records


# -----------------------------------------------------------------------
# B/D Sheet Parsing (Single Row Only)
# -----------------------------------------------------------------------
def parse_bd_row(raw_text):
    """
    Expects exactly one row for B/D.
    We'll parse:
      - MAWB: col F (5)
      - Pieces: col L (11)
      - PMC: col P (15)
      - hold: col U (20)
      - last_mile: col V (21)

    Then we check columns W–Z for normal carriers:
      YUN2 => col W (22)
      UPS  => col X (23)
      UNI  => col Y (24)
      YWE  => col Z (25)

    If column AA (26) is not empty => USPS override:
      carriers = [("800-", ""), ("807-", ""), ("808-", "")]

    This means we skip columns W–Z in that scenario.
    """
    cols = raw_text.strip().split("\t")
    if len(cols) < 27:  # need at least up to index 26 (AA)
        return {}


    hold_str = safe_get(cols, 20).strip()
    # 1) If hold_str is purely numeric, append " HOLD"
    #    e.g. "2" => "2 HOLD"
    if hold_str.isdigit():
        hold_str = hold_str + " HOLD"

    record = {
        "mawb": safe_get(cols, 5).strip(),
        "pieces": safe_get(cols, 11).strip(),
        "pmc": safe_get(cols, 15).strip(),
        "hold": hold_str,
        "last_mile": safe_get(cols, 21).strip(),
        "carriers": []
    }
    if not record["mawb"]:
        # no valid row
        return {}

    # Check if USPS-CO column (AA => index 26) is non-empty
    usps_val = safe_get(cols, 26).strip()
    if usps_val:
        # If there's anything in column AA => override with USPS logic
        # Up to three lines, all with empty ctn counts
        record["carriers"] = [
            ("800-", ""),
            ("807-", ""),
            ("808-", "")
        ]
    else:
        # Normal approach: check columns W–Z for YUN2, UPS, UNI, YWE
        carrier_cols = {
            "YUN2": 22,  # W
            "UPS":  23,  # X
            "UNI":  24,  # Y
            "YWE":  25,  # Z
        }
        found_carriers = []
        for name, idx in carrier_cols.items():
            ctn = safe_get(cols, idx).strip()
            if ctn:
                found_carriers.append((name, ctn))

        # Keep up to 3
        found_carriers = found_carriers[:3]
        record["carriers"] = found_carriers

    return record