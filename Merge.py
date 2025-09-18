import os
import json

def build_json_from_txt(folder_path):
    result = {}

    for filename in os.listdir(folder_path):
        if filename.startswith("datafields_") and filename.endswith(".txt"):
            # Extract key name
            key = filename.replace("datafields_", "").replace(".txt", "")

            # Read lines and remove duplicates while preserving order
            seen = set()
            lines = []
            with open(os.path.join(folder_path, filename), "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and line not in seen:
                        seen.add(line)
                        lines.append(line)

            result[key] = lines

    return result


if __name__ == "__main__":
    folder = "./your_folder_path"  # change to your folder path
    data_json = build_json_from_txt(folder)

    # Save JSON
    with open("output.json", "w", encoding="utf-8") as out:
        json.dump(data_json, out, indent=4, ensure_ascii=False)

    print("✅ JSON created successfully: output.json")
