import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook

# ---------- CONFIGURATION ----------
BASE_PATH = r"T:/a/logical"   # 👈 change this to your base folder path
OUTPUT_EXCEL = "gts_tsq_report.xlsx"
# ----------------------------------

# Initialize Excel workbook and sheets
workbook = Workbook()
sheet_main = workbook.active
sheet_main.title = "Mapping"
sheet_error = workbook.create_sheet("Errors")

# Headers
sheet_main.append(["filename", "subgts", "tsq"])
sheet_error.append(["parent_gts", "lkpath", "missing_part", "checked_path"])

visited_files = set()  # To prevent infinite recursion


# ---------- Function: Resolve lkpath ----------
def resolve_lkpath(base_path, lkpath):
    """
    Resolves lkpath step by step as per rules.
    Returns (found_path, found_type, missing_part)
    found_type = 'gts' | 'tsq' | 'error'
    """
    parts = lkpath.split('/')
    found_base = None
    checked_path = ""
    missing_part = None

    # Step-by-step check (a/b/c/d -> check a, then b, then c, etc.)
    for i in range(len(parts)):
        temp_path = os.path.join(base_path, *parts[:i + 1])
        if os.path.exists(temp_path):
            found_base = temp_path
            checked_path = temp_path
        else:
            continue

    # If no base folder found at all
    if not found_base:
        return (os.path.join(base_path, lkpath), "error", parts[-1])

    # Check final file existence (.gts or .tsq)
    final_candidate = os.path.join(found_base, parts[-1])
    if os.path.exists(final_candidate + ".gts"):
        return (final_candidate + ".gts", "gts", None)
    elif os.path.exists(final_candidate + ".tsq"):
        return (final_candidate + ".tsq", "tsq", None)
    elif os.path.exists(final_candidate):
        # directory exists but not file
        return (final_candidate, "error", parts[-1])
    else:
        # final part missing
        return (final_candidate, "error", parts[-1])


# ---------- Function: Process each file ----------
def process_file(file_path, parent_name):
    """
    Parses XML file, extracts <TestItem> lkpaths,
    resolves them, and logs results.
    """
    if file_path in visited_files:
        return
    visited_files.add(file_path)

    gts_list = []
    tsq_list = []

    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
    except Exception as e:
        sheet_error.append([parent_name, "parse_error", str(e), file_path])
        return

    for test_item in root.findall(".//TestItem"):
        lkpath = test_item.get("lkpath")
        if not lkpath:
            continue

        found_path, found_type, missing_part = resolve_lkpath(BASE_PATH, lkpath)

        if found_type == "gts":
            subgts_name = os.path.splitext(os.path.basename(found_path))[0]
            gts_list.append(subgts_name)
            process_file(found_path, subgts_name)  # recursive call
        elif found_type == "tsq":
            tsq_name = os.path.splitext(os.path.basename(found_path))[0]
            tsq_list.append(tsq_name)
        else:
            sheet_error.append([parent_name, lkpath, missing_part, found_path])

    # Write result to main sheet
    sheet_main.append([
        parent_name,
        ",".join(gts_list) if gts_list else "-",
        ",".join(tsq_list) if tsq_list else "-"
    ])


# ---------- MAIN EXECUTION ----------
def main():
    print("🔍 Scanning directory for .gts and .tsq files...")
    for root_dir, _, files in os.walk(BASE_PATH):
        for file in files:
            if file.endswith(".gts") or file.endswith(".tsq"):
                file_path = os.path.join(root_dir, file)
                filename = os.path.splitext(file)[0]
                process_file(file_path, filename)

    workbook.save(OUTPUT_EXCEL)
    print(f"✅ Report generated successfully: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
def resolve_lkpath(base_path, lkpath):
    """
    Resolve lkpath according to rules:
      - Given lkpath parts = [p0, p1, ..., pn]
      - Check base/p0, if exists then final = base/p0/p1/.../pn => check final.gts / final.tsq
      - If base/p0 does NOT exist, check base/p1, if exists then final = base/p1/p2/.../pn => check final.gts / final.tsq
      - Continue this way. If some base/px exists, we always append remaining parts AFTER px to form final path.
      - If no base/px exists for any px then it's an error (no candidate found).
    Returns tuple: (final_candidate_path, found_type, missing_part)
      - found_type: "gts" | "tsq" | "error"
      - missing_part: last missing path segment (for error reporting) or None
    """
    parts = [p for p in lkpath.split('/') if p]  # remove empty parts
    if not parts:
        return (os.path.join(base_path, lkpath), "error", "empty_lkpath")

    # Try each segment as candidate root under base_path
    for i, seg in enumerate(parts):
        candidate_root = os.path.join(base_path, seg)
        if os.path.exists(candidate_root):
            # We found a starting segment at index i.
            # Build the full path from this found point to the end of parts
            remaining = parts[i+1:]  # parts after the found segment
            # final path starts at candidate_root, then append remaining parts
            final_path_no_ext = os.path.join(candidate_root, *remaining) if remaining else candidate_root

            # Check for .gts or .tsq files (file names expected to be the final segment base name)
            gts_path = final_path_no_ext + ".gts"
            tsq_path = final_path_no_ext + ".tsq"

            if os.path.exists(gts_path):
                return (gts_path, "gts", None)
            if os.path.exists(tsq_path):
                return (tsq_path, "tsq", None)

            # If neither file exists, this is an error: final candidate missing the extension
            # We report the last segment that couldn't be resolved into a .gts/.tsq file
            missing_part = os.path.basename(final_path_no_ext)
            return (final_path_no_ext, "error", missing_part)

    # If loop completes, none of the segments were found directly under base_path
    # This is an error — nothing matched as a candidate root
    # Report the first segment as the initial clue or the whole lkpath as missing
    return (os.path.join(base_path, *parts), "error", parts[-1])
