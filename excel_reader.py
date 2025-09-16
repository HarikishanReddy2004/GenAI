import re
from collections import OrderedDict
import json
import pandas as pd

def count_level(s: str) -> int:
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
    rows: list of tuples (element_raw, type_raw) in sheet order.
    Returns:
      - mapping: OrderedDict { type_token: OrderedDict(childName -> None or list) }
      - all_leaves: ordered list of leaf strings
    """
    n = len(rows)
    i = 0
    result = OrderedDict()
    all_leaves = []

    def parse_children(start_idx, base_level):
        """Parse subtree starting at start_idx where children are level base_level+1."""
        children = []
        idx = start_idx
        while idx < n:
            elem_raw, _ = rows[idx]
            lvl = count_level(elem_raw)
            if lvl <= base_level:
                break
            if lvl == base_level + 1:
                name = extract_after_colon(elem_raw)
                if idx + 1 < n and count_level(rows[idx+1][0]) > lvl:
                    nested_list, new_idx = parse_children(idx + 1, lvl)
                    children.append({name: nested_list})
                    idx = new_idx
                else:
                    children.append(name)
                    idx += 1
            else:
                # malformed deeper line without parent at expected level; skip it safely
                idx += 1
        return children, idx

    while i < n:
        elem_raw, type_raw = rows[i]
        lvl = count_level(elem_raw)
        if lvl != 0:
            i += 1
            continue
        top_elem = extract_after_colon(elem_raw)
        top_type = extract_after_colon(type_raw) if type_raw else top_elem

        # If next row has arrows → subtree; else standalone
        if i + 1 < n and count_level(rows[i+1][0]) > 0:
            children_list, next_i = parse_children(i+1, 0)
            od = OrderedDict()
            for ch in children_list:
                if isinstance(ch, str):
                    od[ch] = None
                elif isinstance(ch, dict):
                    for k, v in ch.items():
                        od[k] = v
            result[top_type] = od

            # collect leaves
            def collect_leaves_from_list(lst):
                for item in lst:
                    if isinstance(item, str):
                        if item not in all_leaves:
                            all_leaves.append(item)
                    elif isinstance(item, dict):
                        for nk, nv in item.items():
                            collect_leaves_from_list(nv)
            collect_leaves_from_list(children_list)
            i = next_i
        else:
            result[top_type] = top_elem
            if top_elem not in all_leaves:
                all_leaves.append(top_elem)
            i += 1

    return result, all_leaves

def mapping_to_jsonable(mapping):
    out = {}
    for top, children in mapping.items():
        if isinstance(children, str):
            out[top] = children
        else:
            conv = {}
            for k, v in children.items():
                conv[k] = v
            out[top] = conv
    return out

# ---------------- Excel Integration ----------------
def process_excel(file_path):
    # Open the "Message Response" sheet
    df = pd.read_excel(file_path, sheet_name="Message Response", dtype=str)

    # Keep only the two needed columns
    df = df[["Response Element Name", "Type"]].fillna("")

    # Convert to list of tuples
    rows = list(zip(df["Response Element Name"], df["Type"]))

    # Parse
    mapping, leaves = parse_rows(rows)

    return mapping, leaves

# ---------------- Example Usage ----------------
if __name__ == "__main__":
    excel_file = "account_list.xlsx"   # your input file
    mapping, leaves = process_excel(excel_file)

    # Save hierarchical mapping
    with open("final_mapping.json", "w") as f:
        json.dump(mapping_to_jsonable(mapping), f, indent=2)

    # Save flat leaf list
    with open("final_leaves.json", "w") as f:
        json.dump(leaves, f, indent=2)

    print("Mapping and leaves extracted successfully!")

# import pandas as pd

# def process_excel(file_path):
#     # Read starting from row 2 (Excel row 3) and only 2nd & 3rd columns
#     df = pd.read_excel(
#         file_path,
#         sheet_name="Message Response",
#         skiprows=2,      # skip first 2 rows (row 0 and row 1)
#         usecols=[1, 2],  # 2nd and 3rd cols (B, C in Excel)
#         dtype=str
#     ).fillna("")

#     # Rename for clarity
#     df.columns = ["Response Element Name", "Type"]

#     # Convert to list of tuples
#     rows = list(zip(df["Response Element Name"], df["Type"]))

#     # Call your existing parsing logic
#     mapping, leaves = parse_rows(rows)

#     return mapping, leaves

