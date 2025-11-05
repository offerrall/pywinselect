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

files = get_selected()
if files:
    print(f"Selected: {files[0]}")

only_files = get_selected(filter_type="files")
only_folders = get_selected(filter_type="folders")
```

## API

### `get_selected(filter_type=None) -> list[str]`

Returns list of absolute paths to selected items.

**Parameters:**
- `filter_type` (optional): Filter results by type
  - `None` - Return both files and folders (default)
  - `"files"` - Return only files
  - `"folders"` - Return only folders

**Returns:** 
- `list[str]` - Paths to selected files/folders
- `[]` - Empty list if nothing is selected

**Examples:**
```python
>>> get_selected()
['C:\\Users\\John\\file.txt', 'C:\\Users\\John\\Documents']

>>> get_selected(filter_type="files")
['C:\\Users\\John\\file.txt']

>>> get_selected(filter_type="folders")
['C:\\Users\\John\\Documents']
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