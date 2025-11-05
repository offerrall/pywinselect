from pywinselect import get_selected


try:
    print("=== All selected items ===")
    all_items = get_selected()
    if all_items:
        for item in all_items:
            print(f"  {item}")
    else:
        print("  Nothing selected")
except Exception as e:
    print(f"Error from get_selected: {e}")

print("\n=== Only files ===")
only_files = get_selected(filter_type="files")
if only_files:
    for file in only_files:
        print(f"  {file}")
else:
    print("  No files selected")

print("\n=== Only folders ===")
only_folders = get_selected(filter_type="folders")
if only_folders:
    for folder in only_folders:
        print(f"  {folder}")
else:
    print("  No folders selected")

input("\nPress Enter to exit...")