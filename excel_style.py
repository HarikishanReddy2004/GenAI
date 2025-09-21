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


import re
import json
import pandas as pd

# ---------------- Utilities ----------------

def count_level(s: str) -> int:
    """Count leading > characters (indentation level)."""
    if s is None or not str(s).strip():
        return 0
    m = re.match(r'^\s*(>+)', str(s))
    return len(m.group(1)) if m else 0

def clean_name(s: str) -> str:
    """Remove > and namespace, keep only name after last ':'."""
    if s is None:
        return ""
    t = str(s).strip()
    t = re.sub(r'^\s*>+\s*', '', t)   # remove leading > 
    if ':' in t:
        t = t.split(':')[-1]
    return t.strip().lower()

def clean_type(t: str) -> str:
    """Clean type: only take part after last ':' (or None if empty)."""
    if t is None or not str(t).strip():
        return None
    tt = str(t).strip()
    if ':' in tt:
        tt = tt.split(':')[-1]
    return tt.strip().lower()

# ---------------- Core structure builder ----------------

def build_structure(rows, start=0, level=0):
    result = {}
    i = start
    n = len(rows)

    while i < n:
        elem_raw, _ = rows[i]
        lvl = count_level(elem_raw)

        if lvl < level:
            break

        if lvl == level:
            name = clean_name(elem_raw)
            next_lvl = count_level(rows[i+1][0]) if (i+1) < n else -1

            if next_lvl > level:
                children, new_i = build_structure(rows, i+1, level+1)
                result[name] = children
                i = new_i
                continue
            else:
                result[name] = []
        i += 1

    return result, i

def process_dataframe(df):
    # remove rows where both columns are empty
    df = df.dropna(how="all")
    df = df[(df["Response Element Name"].astype(str).str.strip() != "") | 
            (df["Type"].astype(str).str.strip() != "")]
    
    rows = list(zip(df["Response Element Name"], df["Type"]))
    
    # mapping
    mapping = {}
    for elem_raw, type_raw in rows:
        key = clean_name(elem_raw)
        mapping[key] = clean_type(type_raw)

    # structure
    structure, _ = build_structure(rows, start=0, level=0)
    return mapping, structure

# ---------------- Extract leaf data fields ----------------

def extract_datafields(structure):
    """Return list of all keys that map to [] (leaf data fields)."""
    fields = []

    def dfs(node):
        for k, v in node.items():
            if v == []:
                fields.append(k)
            elif isinstance(v, dict):
                dfs(v)

    dfs(structure)
    return fields

# ---------------- Batch processing ----------------

def process_files(excel_files, mapping_files, datafield_files):
    for xlsx, jfile, tfile in zip(excel_files, mapping_files, datafield_files):
        print(f"Processing {xlsx} → {jfile}, {tfile}")

        df = pd.read_excel(xlsx, sheet_name=0, usecols=[0,1], header=None)
        df.columns = ["Response Element Name", "Type"]

        mapping, structure = process_dataframe(df)
        datafields = extract_datafields(structure)

        # dump mapping.json
        with open(jfile, "w") as f:
            json.dump(mapping, f, indent=2)

        # dump datafields.txt
        with open(tfile, "w") as f:
            f.write("\n".join(datafields))

        print(f"Done: {xlsx}")

# ---------------- Example Usage ----------------
if __name__ == "__main__":
    a = ["abc.xlsx", "def.xlsx"]      # Excel input list
    b = ["hai.json", "hello.json"]    # Mapping JSON output list
    c = ["hai.txt", "hello.txt"]      # Datafields TXT output list

    process_files(a, b, c)


import re
import json
import pandas as pd
from typing import List, Tuple, Dict, Any

# ---------- Utilities ----------

def count_level(s: str) -> int:
    """Count leading '>' or '/' characters (indentation level)."""
    if s is None:
        return 0
    s = str(s)
    m = re.match(r'^\s*([>/]+)', s)
    return len(m.group(1)) if m else 0

def clean_name(s: str) -> str:
    """
    Remove leading '>' or '/' and keep the text AFTER the last ':' (if present).
    Return lowercased trimmed name.
    """
    if s is None:
        return ""
    t = str(s).strip()
    t = re.sub(r'^\s*[>/]+\s*', '', t)   # strip leading arrows/slashes
    if ':' in t:
        t = t.split(':')[-1]
    return t.strip().lower()

def clean_type(t: str) -> str:
    """Return type (right of colon) or None if empty."""
    if t is None or str(t).strip() == "":
        return None
    tt = str(t).strip()
    if ':' in tt:
        tt = tt.split(':')[-1]
    return tt.strip().lower()

# ---------- Structure builder (recursive) ----------

def parse_structure(rows: List[Tuple[str, str]], start: int = 0, base_level: int = 0) -> Tuple[Dict[str, Any], int]:
    """
    Parse rows into nested dict where leaves are empty lists.
    rows: list of tuples (element_raw, type_raw) in original order.
    Returns: (mapping_dict_at_this_level, next_index_to_process)
    """
    n = len(rows)
    i = start
    result: Dict[str, Any] = {}

    while i < n:
        elem_raw, _ = rows[i]
        lvl = count_level(elem_raw)

        # If we've gone up to a previous level, return to caller
        if lvl < base_level:
            break

        if lvl == base_level:
            name = clean_name(elem_raw)

            # Lookahead to decide if this node has children
            next_lvl = count_level(rows[i+1][0]) if (i + 1) < n else -1

            if next_lvl > base_level:
                # parse children at next level
                children_dict, next_i = parse_structure(rows, i + 1, base_level + 1)
                result[name] = children_dict
                i = next_i
                continue
            else:
                # no children -> leaf (empty list)
                result[name] = []
        # else lvl > base_level should not happen (handled by recursion)
        i += 1

    return result, i

# ---------- Mapping builder (element -> type) ----------

def build_mapping(rows: List[Tuple[str, str]]) -> Dict[str, Any]:
    mapping: Dict[str, Any] = {}
    for elem_raw, type_raw in rows:
        if elem_raw is None or str(elem_raw).strip() == "":
            continue
        name = clean_name(elem_raw)
        typ = clean_type(type_raw)
        mapping[name] = typ if typ is not None else None
    return mapping

# ---------- Main / Example ----------

if __name__ == "__main__":
    # Example you provided (3 rows)
    data = [
        ["vdvbhbd:a", "xsd:string"],
        ["vd shdk:b", "djbvkj:btype"],
        [">dvhbdhvdh:c", "xsd:string"]
    ]

    rows = data  # if you read from Excel, make rows = df[["Response Element Name","Type"]].values.tolist()

    # Build mapping (element -> type)
    mapping_json = build_mapping(rows)

    # Build hierarchical structure from Response Element Name only
    structure_dict, _ = parse_structure(rows, start=0, base_level=0)

    # Print results
    print("Mapping (element -> type):")
    print(json.dumps(mapping_json, indent=2))
    print("\nStructure:")
    print(json.dumps(structure_dict, indent=2))

    # Expected:
    # mapping_json -> {"a":"string","b":"btype","c":"string"}
    # structure_dict -> {"a": [], "b": {"c": []}}
