import re
import json
import pandas as pd
from collections import OrderedDict

# ---------------- Helper Functions ----------------

def count_level(s: str) -> int:
    """Count '>' depth from element column."""
    if s is None or not str(s).strip():
        return 0
    m = re.match(r'^\s*(>+)', str(s))
    return len(m.group(1)) if m else 0

def extract_after_colon(s: str) -> str:
    """Strip leading '>' and return text after the last ':' (or whole text if no colon)."""
    if s is None:
        return ''
    s = str(s).strip()
    s = re.sub(r'^\s*>+\s*', '', s)
    if ':' in s:
        return s.split(':')[-1].strip()
    return s.strip()

def parse_rows(rows):
    """
    rows: list of tuples (element_raw, type_raw)
    Returns:
      mapping (OrderedDict) - hierarchical structure
      all_leaves (list) - flattened leaves
    """
    n = len(rows)
    i = 0
    result = OrderedDict()
    all_leaves = []

    def parse_children(start_idx, base_level):
        children = []
        idx = start_idx
        while idx < n:
            elem_raw, _ = rows[idx]
            lvl = count_level(elem_raw)
            if lvl <= base_level:
                break
            if lvl == base_level + 1:
                name = extract_after_colon(elem_raw)
                # has nested children
                if idx + 1 < n and count_level(rows[idx+1][0]) > lvl:
                    nested_list, new_idx = parse_children(idx + 1, lvl)
                    children.append({name: nested_list})
                    idx = new_idx
                else:
                    children.append(name)
                    if name not in all_leaves:
                        all_leaves.append(name)
                    idx += 1
            else:
                # malformed deeper line → skip
                idx += 1
        return children, idx

    while i < n:
        elem_raw, type_raw = rows[i]
        lvl = count_level(elem_raw)
        if lvl != 0:  # only process top-level (no arrows)
            i += 1
            continue

        top_elem = extract_after_colon(elem_raw)
        top_type = extract_after_colon(type_raw) if type_raw else top_elem

        if i + 1 < n and count_level(rows[i+1][0]) > 0:
            children_list, next_i = parse_children(i + 1, 0)
            od = OrderedDict()
            for ch in children_list:
                if isinstance(ch, str):
                    od[ch] = None  # we will clean later
                elif isinstance(ch, dict):
                    for k, v in ch.items():
                        od[k] = v
            result[top_type] = od
            i = next_i
        else:
            result[top_type] = None
            if top_elem not in all_leaves:
                all_leaves.append(top_elem)
            i += 1

    return result, all_leaves

def mapping_to_jsonable(mapping):
    """Convert OrderedDict structure into JSON-able with correct rules (no nulls)."""
    out = OrderedDict()
    for top, children in mapping.items():
        if children is None:
            # leaf at top
            out[top] = None
        elif isinstance(children, OrderedDict):
            conv = OrderedDict()
            for k, v in children.items():
                if v is None:
                    conv[k] = None  # we keep just the key in JSON
                else:
                    conv[k] = v
            out[top] = conv
        else:
            out[top] = children
    return out

# ---------------- Main Excel Processing ----------------

def process_excel(file_path, sheet_name="Message Response"):
    # Read starting from row 3, cols B & C
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        skiprows=2,
        usecols=[1, 2],
        dtype=str
    ).fillna("")
    df.columns = ["Element", "Type"]

    rows = list(zip(df["Element"], df["Type"]))
    mapping, leaves = parse_rows(rows)
    return mapping, leaves

# ---------------- Example Usage ----------------
if __name__ == "__main__":
    excel_file = "account_list.xlsx"
    mapping, leaves = process_excel(excel_file)

    # Save hierarchical mapping
    with open("final_mapping.json", "w") as f:
        json.dump(mapping_to_jsonable(mapping), f, indent=2)

    # Save flat leaf list
    with open("final_leaves.json", "w") as f:
        json.dump(leaves, f, indent=2)

    print("Mapping and leaves extracted successfully!")
