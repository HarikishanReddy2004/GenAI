import re
import json
import os
import pandas as pd

# ---------------- Utilities ----------------

def count_level(s: str) -> int:
    """Count leading '>' characters (indentation level)."""
    if s is None or not str(s).strip():
        return 0
    m = re.match(r'^\s*(>+)', str(s))
    return len(m.group(1)) if m else 0

def clean_name(s: str) -> str:
    """Strip leading >, take text after last ':' and lowercase."""
    if s is None:
        return ""
    t = str(s).strip()
    t = re.sub(r'^\s*>+\s*', '', t)        # remove leading arrows
    if ':' in t:
        t = t.split(':')[-1]
    return t.strip().lower()

def clean_type(t: str) -> str:
    """Take text after last ':' in type and lowercase (or None)."""
    if t is None or not str(t).strip():
        return None
    tt = str(t).strip()
    if ':' in tt:
        tt = tt.split(':')[-1]
    return tt.strip().lower()

# ---------------- Core parser ----------------

def build_structure(rows, start=0, level=0):
    result = []
    i = start
    n = len(rows)

    while i < n:
        elem_raw, type_raw = rows[i]
        lvl = count_level(elem_raw)

        if lvl < level:
            break

        if lvl == level:
            name = clean_name(elem_raw)
            next_lvl = count_level(rows[i+1][0]) if (i + 1) < n else -1

            if next_lvl > level:
                key = clean_type(type_raw) or name
                children, new_i = build_structure(rows, i+1, level+1)
                result.append({key: children})
                i = new_i
                continue
            else:
                # Leaf node
                result.append(name)
        i += 1

    return result, i

def process_excel(file_path):
    """Read Excel, clean empty rows, return structured data."""
    _, ext = os.path.splitext(file_path)
    engine = None
    if ext.lower() == ".xls":
        engine = "xlrd"  # requires xlrd==1.2.0 for .xls

    df = pd.read_excel(
        file_path,
        sheet_name="Message Response",
        skiprows=2,      # start at Excel row 3
        usecols=[1, 2],  # B and C columns
        dtype=str,
        engine=engine
    )

    # Drop completely empty rows
    df = df.dropna(how="all").fillna("")

    df.columns = ["Response Element Name", "Type"]
    rows = list(zip(df["Response Element Name"], df["Type"]))

    # Clean empty element names
    rows = [(a, b) for a, b in rows if str(a).strip()]

    structure, _ = build_structure(rows, start=0, level=0)

    # ensure outer structure is a dict
    if isinstance(structure, list) and len(structure) == 1 and isinstance(structure[0], dict):
        return structure[0]
    return {"root": structure}

# ---------------- Leaf extractor ----------------

def extract_leaves(node, leaves=None):
    """Recursively extract all leaf strings from the nested structure."""
    if leaves is None:
        leaves = []

    if isinstance(node, str):
        leaves.append(node)
    elif isinstance(node, dict):
        for v in node.values():
            extract_leaves(v, leaves)
    elif isinstance(node, list):
        for item in node:
            extract_leaves(item, leaves)

    return leaves

# ---------------- Runner ----------------

if __name__ == "__main__":
    excel_file = "account_list.xlsx"   # change to your file
    final_structure = process_excel(excel_file)

    # Save JSON
    with open("final_structure.json", "w", encoding="utf-8") as f:
        json.dump(final_structure, f, indent=2, ensure_ascii=False)

    # Extract leaves
    leaves = extract_leaves(final_structure)

    # Save leaves
    with open("leaf_nodes.txt", "w", encoding="utf-8") as f:
        for leaf in leaves:
            f.write(leaf + "\n")

    print("=== Final Structure ===")
    print(json.dumps(final_structure, indent=2))

    print("\n=== Extracted Leaves ===")
    print(leaves)



import re
import json
import pandas as pd

# ---------------- Utilities ----------------

def count_level(s: str) -> int:
    """Count leading '>' characters (indentation level)."""
    if s is None or not str(s).strip():
        return 0
    m = re.match(r'^\s*(>+)', str(s))
    return len(m.group(1)) if m else 0

def clean_name(s: str) -> str:
    """Strip leading >, take text after last ':' and lowercase."""
    if s is None:
        return ""
    t = str(s).strip()
    t = re.sub(r'^\s*>+\s*', '', t)        # remove leading arrows
    if ':' in t:
        t = t.split(':')[-1]
    return t.strip().lower()

def clean_type(t: str) -> str:
    """Take text after last ':' in type and lowercase (or None)."""
    if t is None or not str(t).strip():
        return None
    tt = str(t).strip()
    if ':' in tt:
        tt = tt.split(':')[-1]
    return tt.strip().lower()

# ---------------- Core parser ----------------

def build_structure(rows, start=0, level=0):
    result = []
    i = start
    n = len(rows)

    while i < n:
        elem_raw, type_raw = rows[i]
        lvl = count_level(elem_raw)

        if lvl < level:
            break

        if lvl == level:
            name = clean_name(elem_raw)
            next_lvl = count_level(rows[i+1][0]) if (i + 1) < n else -1

            if next_lvl > level:
                key = clean_type(type_raw) or name
                children, new_i = build_structure(rows, i+1, level+1)
                result.append({key: children})
                i = new_i
                continue
            else:
                # Leaf node
                result.append(name)
        i += 1

    return result, i

def process_dataframe(df):
    rows = list(zip(df["Response Element Name"], df["Type"]))
    structure, _ = build_structure(rows, start=0, level=0)

    # ensure outer structure is a dict (for JSON)
    if isinstance(structure, list) and len(structure) == 1 and isinstance(structure[0], dict):
        return structure[0]
    return {"root": structure}

# ---------------- Leaf extractor ----------------

def extract_leaves(node, leaves=None):
    """Recursively extract all leaf strings from the nested structure."""
    if leaves is None:
        leaves = []

    if isinstance(node, str):
        leaves.append(node)
    elif isinstance(node, dict):
        for v in node.values():
            extract_leaves(v, leaves)
    elif isinstance(node, list):
        for item in node:
            extract_leaves(item, leaves)

    return leaves

# ---------------- Test with provided data ----------------

data = [
    ["customerToolUsage", "CustomerToolUsage_Type"],
    [">channel:channelKey", "channel:ChannelKey_Type"],
    [">channel:deliveryChannelCode", "base:DeliveryChannelCode_Enum"],
    ["channel:computerApplicationCode", "base:ComputerApplicationCode_Type"],
    [">prflQstn:toolCode", "xs:string"],
    [">prflQstn:usageById", "xs:string"],
    [">prflQstn:usageByIdType", "xs:string"],
    [">teamMember:teamMemberName", "base:EntityName_Type"],
    [">>format", "base:NameFormat_Enum"],
    [">>nameComponents", "base:NameComponents_Type"],
    [">>>namePrefix", "xs:string"],
    [">>>firstName", "xs:string"],
    [">>>middleName", "xs:string"],
    [">>>lastName", "xs:string"],
    [">>>nameSuffix", "xs:string"],
    [">>fullName", "base:NameLineList_Type"],
    [">>>nameLine", "xs:string"],
    [">phoneticName", "xs:string"],
    ["prflQstn:updateDate", "xs:date"],
]
df = pd.DataFrame(data, columns=["Response Element Name", "Type"])

# Step 1: build structure
final_structure = process_dataframe(df)

# Step 2: Save JSON
with open("final_structure.json", "w", encoding="utf-8") as f:
    json.dump(final_structure, f, indent=2, ensure_ascii=False)

# Step 3: Extract leaves
leaves = extract_leaves(final_structure)

# Step 4: Save leaves to TXT
with open("leaf_nodes.txt", "w", encoding="utf-8") as f:
    for leaf in leaves:
        f.write(leaf + "\n")

print("=== Final Structure ===")
print(json.dumps(final_structure, indent=2))

print("\n=== Extracted Leaves ===")
print(leaves)

