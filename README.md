# pywinselect

Get selected files and folders in Windows.

## Installation
```bash
pip install git+https://github.com/offerrall/pywinselect
```

## What it does

`pywinselect` detects which files or folders you have selected in Windows, regardless of where they are.

It works in:
- **File Explorer** - Any open Explorer window
- **Desktop** - Files and folders selected on your desktop

Returns absolute paths to everything selected. If nothing is selected, returns an empty list.

## Why use it

When building automation tools, you often need to know what the user has selected. Without this library, you'd have to:

- Write 100+ lines of win32 API code
- Handle File Explorer and Desktop separately  
- Deal with clipboard backup/restore
- Debug why it doesn't work on Desktop

`pywinselect` solves this in one line.

## Use cases

- **Stream Deck scripts** - Create buttons that act on selected files
- **Keyboard macros** - Automate repetitive file operations
- **Custom context menus** - Build tools that work with selections
- **Productivity apps** - Quick actions on currently selected items
- **Batch processing** - Process whatever the user has selected

## Usage
```python
from pywinselect import get_selected

# Get all selected items (Desktop + Explorer)
files = get_selected()

# Filter by type
only_files = get_selected(filter_type="files")
only_folders = get_selected(filter_type="folders")

# Get selection from specific location
desktop_only = get_selected(only_desktop=True)
explorer_only = get_selected(only_explorer=True)
```

## API

### `get_selected(filter_type=None, only_desktop=False, only_explorer=False) -> list[str]`

Returns list of absolute paths to selected items.

**Parameters:**
- `filter_type` (optional): Filter results by type
  - `None` - Return both files and folders (default)
  - `"files"` - Return only files
  - `"folders"` - Return only folders
- `only_desktop` (optional): If `True`, only get selection from Desktop
- `only_explorer` (optional): If `True`, only get selection from File Explorer

**Returns:** 
- `list[str]` - Paths to selected files/folders
- `[]` - Empty list if nothing is selected

**Raises:**
- `ValueError` - If both `only_desktop` and `only_explorer` are `True`

**Examples:**
```python
>>> get_selected()
['C:\\Users\\John\\file.txt', 'C:\\Users\\John\\Documents']

>>> get_selected(filter_type="files")
['C:\\Users\\John\\file.txt']

>>> get_selected(filter_type="folders")
['C:\\Users\\John\\Documents']

>>> get_selected(only_desktop=True)
['C:\\Users\\John\\Desktop\\file.txt']

>>> get_selected(only_explorer=True)
['C:\\Users\\John\\Downloads\\file.txt']

>>> get_selected(only_desktop=True, only_explorer=True)
ValueError: Cannot set both only_desktop and only_explorer to True
```

## How it works

Uses official Windows Shell COM APIs (`IShellView`, `IDataObject`) - the same interfaces Windows Explorer uses internally. 

**Safe by design:**
- Read-only operations
- No keyboard simulation
- No clipboard modifications
- No system state changes

## Requirements

- Python 3.10+
- Windows
- pywin32

## License

MIT