# parse_account_list_final.py
import re
import os
import json
from collections import OrderedDict

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

def _parse_children(rows, start_idx, base_level):
    """
    Parse children starting at start_idx where children have indentation level base_level+1.
    Returns (children_list, next_index).
    children_list uses elements:
      - string => leaf child
      - dict {childName: [ ... ]} => child with nested children (list may contain strings or dicts)
    """
    n = len(rows)
    children = []
    idx = start_idx
    while idx < n:
        elem_raw, _ = rows[idx]
        lvl = count_level(elem_raw)
        if lvl <= base_level:
            break
        if lvl == base_level + 1:
            name = extract_after_colon(elem_raw)
            # if this child has deeper descendants
            if (idx + 1) < n and count_level(rows[idx + 1][0]) > lvl:
                nested, new_idx = _parse_children(rows, idx + 1, lvl)
                children.append({name: nested})
                idx = new_idx
            else:
                children.append(name)
                idx += 1
        else:
            # deeper indentation without proper parent -> skip safely
            idx += 1
    return children, idx

def parse_rows(rows):
    """
    rows: list of tuples (element_raw, type_raw) in sheet order.
    Returns:
      - top_entries: ordered list of ("leaf", name) or ("mapping", key, children_list)
      - leaves: ordered unique list of leaf names
    """
    n = len(rows)
    i = 0
    top_entries = []   # preserve original order: ('leaf', name) or ('mapping', key, children_list)
    seen = set()
    leaves = []

    def collect_leaves_from_items(items):
        for it in items:
            if isinstance(it, str):
                if it not in seen:
                    seen.add(it); leaves.append(it)
            elif isinstance(it, dict):
                for nk, nv in it.items():
                    collect_leaves_from_items(nv)

    while i < n:
        elem_raw, type_raw = rows[i]
        # skip blank element names
        if not elem_raw or not str(elem_raw).strip():
            i += 1
            continue

        lvl = count_level(elem_raw)
        if lvl != 0:
            # mis-indented top-level row: skip (child rows are handled via parent's parse)
            i += 1
            continue

        top_elem = extract_after_colon(elem_raw)
        # Decide based on next row
        if (i + 1) < n:
            next_lvl = count_level(rows[i + 1][0])
        else:
            next_lvl = -1

        # Case: next row also top-level -> current is a standalone leaf
        if next_lvl == 0:
            top_entries.append(('leaf', top_elem))
            if top_elem not in seen:
                seen.add(top_elem); leaves.append(top_elem)
            i += 1
            continue

        # Case: next row is indented -> current's TYPE becomes the mapping key
        if next_lvl > 0:
            # Use type column's right-side as key; fallback to top_elem if type empty
            key = extract_after_colon(type_raw) if (type_raw and str(type_raw).strip()) else top_elem
            children_list, next_i = _parse_children(rows, i + 1, 0)
            top_entries.append(('mapping', key, children_list))
            # collect leaves from this subtree
            collect_leaves_from_items(children_list)
            i = next_i
            continue

        # Case: last row (no next) -> standalone
        top_entries.append(('leaf', top_elem))
        if top_elem not in seen:
            seen.add(top_elem); leaves.append(top_elem)
        i += 1

    return top_entries, leaves

def format_item_compact(item):
    """Return compact string for a child item (string or {k:[..]}). No quotes, no nulls."""
    if isinstance(item, str):
        return item
    elif isinstance(item, dict):
        # single-key dict
        for k, v in item.items():
            inner = ",".join(format_item_compact(x) for x in v)
            return "{" + f"{k}:[{inner}]" + "}"
    return ""

def build_compact_text(top_entries):
    parts = []
    for ent in top_entries:
        if ent[0] == 'leaf':
            parts.append(ent[1])
        else:
            # ('mapping', key, children)
            key = ent[1]
            children = ent[2]
            children_str = ",".join(format_item_compact(ch) for ch in children)
            parts.append("{" + f"{key}:{{" + children_str + "}" + "}}")
    return "{" + ",".join(parts) + "}"

# ---------- Excel reading + main ----------

def process_excel(file_path):
    _, ext = os.path.splitext(file_path)
    engine = None
    if ext.lower() == ".xls":
        engine = "xlrd"  # ensure xlrd==1.2.0 installed for .xls

    df = pd.read_excel(
        file_path,
        sheet_name="Message Response",
        skiprows=2,      # start at Excel row 3 (0-based skiprows)
        usecols=[1, 2],  # B and C columns (Response Element Name, Type)
        dtype=str,
        engine=engine
    ).fillna("")

    df.columns = ["Response Element Name", "Type"]
    rows = list(zip(df["Response Element Name"], df["Type"]))
    top_entries, leaves = parse_rows(rows)
    return top_entries, leaves

if __name__ == "__main__":
    excel_file = "account_list.xlsx"   # change to your file (xls/xlsx)
    top_entries, leaves = process_excel(excel_file)

    # Write compact mapping text exactly in the format you requested (no : null)
    compact_str = build_compact_text(top_entries)
    with open("final_mapping_compact.txt", "w", encoding="utf-8") as f:
        f.write(compact_str)

    # Write final leaves (flat list)
    with open("final_leaves.json", "w", encoding="utf-8") as f:
        json.dump(leaves, f, indent=2)
