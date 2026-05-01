import os

def print_tree(start_path, file, prefix=""):
    try:
        entries = sorted(os.listdir(start_path))
    except PermissionError:
        return

    entries_count = len(entries)

    for index, entry in enumerate(entries):
        full_path = os.path.join(start_path, entry)
        is_last = index == entries_count - 1

        connector = "└── " if is_last else "├── "
        line = prefix + connector + entry
        file.write(line + "\n")

        if os.path.isdir(full_path):
            extension = "    " if is_last else "│   "
            print_tree(full_path, file, prefix + extension)


if __name__ == "__main__":
    ROOT_PATH = r"C:\Users\admin\Documents\test\git1\git\Circuitbuilding"   # ✅ your main folder
    OUTPUT_FILE = "folder_structure.txt"

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        folder_name = os.path.basename(os.path.normpath(ROOT_PATH))
        f.write(folder_name + "/\n")
        print_tree(ROOT_PATH, f)

    print(f"✅ Folder structure saved to: {OUTPUT_FILE}")
