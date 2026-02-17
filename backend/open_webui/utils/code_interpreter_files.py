"""
Code Interpreter File Access Utilities

This module provides utilities for generating file download preamble code
that enables the Jupyter kernel to access user-uploaded files.

Add this file to: backend/open_webui/utils/code_interpreter_files.py
"""

import logging
from typing import List, Dict, Tuple, Optional

log = logging.getLogger(__name__)


def generate_file_access_preamble(
    files: List[Dict], 
    webui_url: str, 
    token: str
) -> Tuple[str, str]:
    """
    Generate Python code preamble that downloads user-uploaded files
    and makes them available via the FILES dictionary.
    
    Args:
        files: List of file metadata dicts with 'id', 'name', 'type' keys
        webui_url: Base URL of the Open WebUI instance
        token: JWT token for authenticating file downloads
        
    Returns:
        Tuple of (preamble_code, cleanup_code) strings
    """
    if not files:
        return "", ""
    
    # Filter to only downloadable files (not images, not collections, etc.)
    downloadable_files = [
        f for f in files 
        if f.get("type") == "file" and f.get("id")
    ]
    
    if not downloadable_files:
        log.debug("No downloadable files found in file list")
        return "", ""
    
    log.info(f"Generating file access preamble for {len(downloadable_files)} files")
    
    # Build the file download entries
    file_entries = []
    for f in downloadable_files:
        file_id = f.get("id", "")
        # Use the display name, falling back to id if not available
        file_name = (
            f.get("name") 
            or f.get("file", {}).get("filename") 
            or f.get("file", {}).get("meta", {}).get("name")
            or file_id
        )
        # Escape quotes in filename for Python string
        file_name_escaped = file_name.replace("\\", "\\\\").replace("'", "\\'")
        file_entries.append(f"    ('{file_id}', '{file_name_escaped}')")
        log.debug(f"File entry: {file_id} -> {file_name}")
    
    files_list = ",\n".join(file_entries)
    
    # Normalize webui_url (remove trailing slash)
    webui_url = webui_url.rstrip("/")
    
    preamble = f'''
# ============================================================
# AUTO-INJECTED FILE ACCESS PREAMBLE
# Downloaded files are available in the FILES dictionary.
# ============================================================
import os as _os
import requests as _requests
import shutil as _shutil

_UPLOAD_DIR = '/tmp/code_interpreter_files'
_os.makedirs(_UPLOAD_DIR, exist_ok=True)

FILES = {{}}
_FILE_DOWNLOAD_TOKEN = '{token}'
_WEBUI_URL = '{webui_url}'

_files_to_download = [
{files_list}
]

for _file_id, _file_name in _files_to_download:
    try:
        _download_url = f"{{_WEBUI_URL}}/api/v1/files/{{_file_id}}/content/jupyter"
        _headers = {{"Authorization": f"Bearer {{_FILE_DOWNLOAD_TOKEN}}"}}
        _response = _requests.get(_download_url, headers=_headers, timeout=60)
        _response.raise_for_status()
        
        # Sanitize filename for filesystem
        _safe_name = "".join(c if c.isalnum() or c in ".-_ " else "_" for c in _file_name)
        _file_path = f"{{_UPLOAD_DIR}}/{{_safe_name}}"
        
        with open(_file_path, 'wb') as _f:
            _f.write(_response.content)
        
        FILES[_file_name] = _file_path
        print(f"[File Loaded] {{_file_name}}")
    except Exception as _e:
        print(f"[File Download Error] {{_file_name}}: {{_e}}")

# Clean up temporary variables (keep FILES and cleanup utilities)
for _tmp_var in [
    "_file_id",
    "_file_name",
    "_download_url",
    "_headers",
    "_response",
    "_safe_name",
    "_file_path",
    "_f",
    "_files_to_download",
    "_FILE_DOWNLOAD_TOKEN",
    "_WEBUI_URL",
]:
    globals().pop(_tmp_var, None)
del _tmp_var

# ============================================================
# USER CODE BEGINS
# ============================================================

'''
    
    # Cleanup code to append AFTER user code
    cleanup = '''

# ============================================================
# AUTO-INJECTED CLEANUP
# ============================================================
try:
    _shutil.rmtree(_UPLOAD_DIR, ignore_errors=True)
except:
    pass
'''
    
    return preamble, cleanup


def get_downloadable_file_count(files: Optional[List[Dict]]) -> int:
    """
    Count how many files in the list are downloadable.
    
    Args:
        files: List of file metadata dicts
        
    Returns:
        Number of downloadable files
    """
    if not files:
        return 0
    
    return len([
        f for f in files 
        if f.get("type") == "file" and f.get("id")
    ])
