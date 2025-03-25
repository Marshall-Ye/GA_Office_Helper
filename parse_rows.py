def parse_rows_from_input():
    """
    Reads multiple lines from stdin until an empty line is encountered.
    Parses each row (tab-separated) and extracts:
      - Column F (index 5)   => MAWB
      - Column J (index 9)   => Flight No
      - Column L (index 11)  => # of Pieces
      - Column N (index 13)  => Weight
    Skips any row missing a valid MAWB.
    Returns a list of dictionaries.
    """
    print("Paste your rows of data (hit Enter on an empty line to finish):")
    lines = []
    while True:
        line = input().strip()
        if not line:
            break
        lines.append(line)

    parsed_records = []

    for row in lines:
        # Split the row by tab
        cols = row.split("\t")

        def safe_get(lst, idx):
            return lst[idx] if idx < len(lst) else ""

        # Extract the fields we need
        mawb = safe_get(cols, 5)  # F column
        flt = safe_get(cols, 9)  # J column
        pieces = safe_get(cols, 11)  # L column
        weight = safe_get(cols, 13)  # N column

        # Skip this row if MAWB is empty
        if not mawb.strip():
            continue

        # Add a dictionary for each valid row
        parsed_records.append({
            "mawb": mawb,
            "flt": flt,
            "pieces": pieces,
            "weight": weight
        })

    return parsed_records
