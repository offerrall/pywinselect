from pywinselect import get_selected


print("=== All selected items (Desktop + Explorer) ===")
all_items = get_selected()
if all_items:
    for item in all_items:
        print(f"  {item}")
else:
    print("  Nothing selected")

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

print("\n=== Only from Desktop ===")
desktop_items = get_selected(only_desktop=True)
if desktop_items:
    for item in desktop_items:
        print(f"  {item}")
else:
    print("  Nothing selected on Desktop")

print("\n=== Only from File Explorer ===")
explorer_items = get_selected(only_explorer=True)
if explorer_items:
    for item in explorer_items:
        print(f"  {item}")
else:
    print("  Nothing selected in Explorer")

print("\n=== Error example: both only_desktop and only_explorer ===")
try:
    get_selected(only_desktop=True, only_explorer=True)
except ValueError as e:
    print(f"  ValueError: {e}")

input("\nPress Enter to exit...")