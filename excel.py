import pandas as pd
import json
import re

def count_level(s: str) -> int:
    """Count indentation level from '>' symbols."""
    if not s or not str(s).strip():
        return 0
    m = re.match(r'^\s*(>+)', str(s))
    return len(m.group(1)) if m else 0

def clean_name(s: str) -> str:
    """Remove > and namespace prefixes like abc:xyz → xyz, lowercase."""
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r'^\s*>+\s*', '', s)   # remove leading >
    if ':' in s:
        s = s.split(':')[-1]
    return s.strip().lower()

def parse_block(rows, start_idx, base_level):
    """Recursively parse children."""
    items = []
    i = start_idx
    while i < len(rows):
        raw_elem, raw_type = rows[i]
        lvl = count_level(raw_elem)
        if lvl <= base_level:
            break

        name = clean_name(raw_elem)

        # if child has children
        if i + 1 < len(rows) and count_level(rows[i+1][0]) > lvl:
            children, new_i = parse_block(rows, i+1, lvl)
            items.append({name: children})
            i = new_i
        else:
            items.append(name)
            i += 1
    return items, i

def parse_rows(rows):
    """Convert rows into nested JSON-like structure with type handling."""
    mapping = {}
    i = 0
    while i < len(rows):
        raw_elem, raw_type = rows[i]
        if not raw_elem.strip():
            i += 1
            continue

        if count_level(raw_elem) != 0:
            i += 1
            continue

        # default name from element
        name = clean_name(raw_elem)

        # check next row level
        if i + 1 < len(rows):
            next_lvl = count_level(rows[i+1][0])
        else:
            next_lvl = -1

        if next_lvl > 0:
            # *** use TYPE column as key ***
            key = clean_name(raw_type) if raw_type.strip() else name
            children, new_i = parse_block(rows, i+1, 0)
            mapping[key] = children
            i = new_i
        else:
            mapping[name] = []
            i += 1
    return mapping

def process_excel(file_path):
    df = pd.read_excel(
        file_path,
        sheet_name="Message Response",
        skiprows=2,
        usecols=[1, 2],  # B=Response Element Name, C=Type
        dtype=str
    ).fillna("")

    rows = list(zip(df.iloc[:,0], df.iloc[:,1]))
    return parse_rows(rows)

if __name__ == "__main__":
    excel_file = "account_list.xlsx"   # change path
    mapping = process_excel(excel_file)

    with open("final_mapping.json", "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=2)

    print("✅ final_mapping.json written")
