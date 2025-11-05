import os
import pythoncom
import win32com.client

try:
    from win32com.shell import shell, shellcon # type: ignore 
except ImportError:
    import win32com.client
    win32com.client.gencache.EnsureModule('{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}', 0, 1, 0)
    from win32com.shell import shell, shellcon # type: ignore

from typing import Literal


def get_selected(filter_type: Literal["files", "folders"] | None = None) -> list[str]:
    """
    Get currently selected files and folders in Windows.
    
    Works in File Explorer and Desktop.
    
    Args:
        filter_type: Optional filter for results.
            - None: Return both files and folders (default)
            - "files": Return only files
            - "folders": Return only folders
    
    Returns:
        List of absolute paths to selected items. Empty list if nothing is selected
        or nothing matches the filter.
        
    Examples:
        >>> from pywinselect import get_selected
        >>> get_selected()
        ['C:\\file.txt', 'C:\\folder']
        >>> get_selected(filter_type="files")
        ['C:\\file.txt']
        >>> get_selected(filter_type="folders")
        ['C:\\folder']
    """
    result = _get_from_explorer()
    
    if not result:
        result = _get_from_desktop()
    
    if not result:
        return []
    
    if filter_type is None:
        return result
    
    if filter_type == "files":
        return [path for path in result if os.path.isfile(path)]
    
    if filter_type == "folders":
        return [path for path in result if os.path.isdir(path)]
    
    return result


def _get_from_explorer() -> list[str] | None:
    """
    Get selection from File Explorer windows.
    
    Returns:
        List of paths if selection found in Explorer, None otherwise.
    """
    try:
        shell_app = win32com.client.Dispatch("Shell.Application")
        for window in shell_app.Windows():
            try:
                items = window.Document.SelectedItems()
                if items.Count > 0:
                    return [item.Path for item in items]
            except:
                pass
    except:
        pass
    return None


def _get_from_desktop() -> list[str] | None:
    """
    Get selection from Desktop using IShellView and IDataObject.
    
    Returns:
        List of paths if selection found on Desktop, None otherwise.
    """
    try:
        CLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
        
        shell_windows = win32com.client.Dispatch(CLSID_ShellWindows)
        
        dispatch = shell_windows.FindWindowSW(
            win32com.client.VARIANT(pythoncom.VT_I4, shellcon.CSIDL_DESKTOP),
            win32com.client.VARIANT(pythoncom.VT_EMPTY, None),
            0x08,
            0,
            0x01,
        )
        
        if not dispatch:
            return None
        
        service_provider = dispatch._oleobj_.QueryInterface(pythoncom.IID_IServiceProvider)
        browser = service_provider.QueryService(
            shell.SID_STopLevelBrowser,
            shell.IID_IShellBrowser
        )
        shell_view = browser.QueryActiveShellView()
        
        data_object = shell_view.GetItemObject(shellcon.SVGIO_SELECTION, pythoncom.IID_IDataObject)
        
        item_array = shell.SHCreateShellItemArrayFromDataObject(data_object)
        
        count = item_array.GetCount()
        
        if count == 0:
            return None
        
        selection = []
        for i in range(count):
            item = item_array.GetItemAt(i)
            path = item.GetDisplayName(shellcon.SIGDN_FILESYSPATH)
            selection.append(path)
        
        return selection if selection else None
        
    except:
        return None