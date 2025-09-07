#!/usr/bin/env python
"""
MCP server exposing Microsoft Word COM operations on Windows.

Tools exposed under namespace 'word.*':
 - word.open(path)
 - word.set_track_changes(on=True)
 - word.insert_comment(anchor, text)
 - word.write_revision(anchor, new_text)
 - word.accept_revision(index)
 - word.reject_revision(index)
 - word.goto(target)
 - word.save_as(path)
 - word.close(save=False)

Anchor can be:
 - {"find": "text to find"}  -> first match in document
 - {"bookmark": "BookmarkName"}
 - {"range": [start, end]}    -> word character positions (1-based)

Serializes access: holds a single Word.Application instance.
"""
import json
import os
import sys
import threading
from typing import Optional, Dict, Any

try:
    from win32com.client import Dispatch
    import pythoncom
except Exception as e:
    print(json.dumps({"error": f"pywin32 not available: {e}"}))
    sys.exit(1)

try:
    # Official MCP Python SDK
    from mcp.server import Server
    from mcp.types import Tool, TextContent
except Exception as e:
    print("This server requires the 'mcp' Python package. Install with: pip install mcp")
    sys.exit(1)


class WordSession:
    def __init__(self):
        self._lock = threading.RLock()
        self.app = None
        self.doc = None
        self.original_path = None

    def _ensure_app(self):
        if self.app is None:
            pythoncom.CoInitialize()
            self.app = Dispatch("Word.Application")
            self.app.Visible = False

    def _find_range(self, anchor: Dict[str, Any]):
        if self.doc is None:
            raise RuntimeError("No document open")
        rng = self.doc.Content
        if not anchor:
            return rng
        if isinstance(anchor, dict):
            if 'bookmark' in anchor:
                try:
                    return self.doc.Bookmarks(anchor['bookmark']).Range
                except Exception:
                    pass
            if 'range' in anchor and isinstance(anchor['range'], (list, tuple)) and len(anchor['range']) == 2:
                start, end = anchor['range']
                return self.doc.Range(Start=int(start), End=int(end))
            if 'find' in anchor:
                text = anchor['find'] or ''
                find = rng.Find
                find.Text = text
                find.MatchCase = False
                find.MatchWholeWord = False
                if find.Execute():
                    return rng
        # fallback to full document
        return rng

    def tool_open(self, path: str) -> str:
        with self._lock:
            self._ensure_app()
            if self.doc is not None:
                self.doc.Close(SaveChanges=False)
                self.doc = None
            self.doc = self.app.Documents.Open(os.path.abspath(path))
            # Never overwrite the originally opened path
            self.original_path = os.path.abspath(path)
            return os.path.abspath(path)

    def tool_set_track(self, on: bool = True) -> bool:
        with self._lock:
            if self.doc is None:
                raise RuntimeError("No document open")
            self.doc.TrackRevisions = bool(on)
            return bool(on)

    def tool_insert_comment(self, anchor: Dict[str, Any], text: str) -> int:
        with self._lock:
            rng = self._find_range(anchor)
            c = self.doc.Comments.Add(rng, text)
            return int(c.Index)

    def tool_write_revision(self, anchor: Dict[str, Any], new_text: str) -> Dict[str, Any]:
        with self._lock:
            rng = self._find_range(anchor)
            # Replace selection to create tracked del+ins if TrackRevisions is on
            rng.Text = new_text
            return {"start": int(rng.Start), "end": int(rng.End), "length": int(rng.End - rng.Start)}

    def tool_accept_revision(self, index: int) -> bool:
        with self._lock:
            revs = self.doc.Revisions
            if 1 <= index <= revs.Count:
                revs.Item(index).Accept()
                return True
            return False

    def tool_reject_revision(self, index: int) -> bool:
        with self._lock:
            revs = self.doc.Revisions
            if 1 <= index <= revs.Count:
                revs.Item(index).Reject()
                return True
            return False

    def tool_goto(self, target: Any) -> Dict[str, int]:
        with self._lock:
            sel = self.app.Selection
            if isinstance(target, int):
                sel.GoTo(What=2, Which=1, Count=target)  # wdGoToSection=2
            elif isinstance(target, str):
                try:
                    rng = self.doc.Bookmarks(target).Range
                    rng.Select()
                except Exception:
                    # try find text
                    rng = self.doc.Content
                    f = rng.Find
                    f.Text = target
                    if f.Execute():
                        rng.Select()
            return {"start": int(sel.Range.Start), "end": int(sel.Range.End)}

    def tool_save_as(self, path: str) -> str:
        with self._lock:
            if self.doc is None:
                raise RuntimeError("No document open")
            dest = os.path.abspath(path)
            if self.original_path and os.path.abspath(dest) == os.path.abspath(self.original_path):
                raise RuntimeError("Refusing to overwrite original document")
            directory = os.path.dirname(dest)
            os.makedirs(directory, exist_ok=True)
            self.doc.SaveAs2(dest)
            return dest

    def tool_close(self, save: bool = False) -> bool:
        with self._lock:
            if self.doc is not None:
                self.doc.Close(SaveChanges=bool(save))
                self.doc = None
            if self.app is not None:
                self.app.Quit()
                self.app = None
            pythoncom.CoUninitialize()
            return True


session = WordSession()
server = Server("word-com")


@server.tool()
def open(path: str) -> str:
    """Open a .docx file for editing (copy decoupee)."""
    return session.tool_open(path)


@server.tool()
def set_track_changes(on: bool = True) -> bool:
    """Enable/disable track changes on the active document."""
    return session.tool_set_track(on)


@server.tool()
def insert_comment(anchor: Dict[str, Any], text: str) -> int:
    """Insert a Word comment at the given anchor."""
    return session.tool_insert_comment(anchor, text)


@server.tool()
def write_revision(anchor: Dict[str, Any], new_text: str) -> Dict[str, Any]:
    """Write revised text at the given anchor (tracked)."""
    return session.tool_write_revision(anchor, new_text)


@server.tool()
def accept_revision(index: int) -> bool:
    """Accept revision by 1-based index in the Revisions collection."""
    return session.tool_accept_revision(index)


@server.tool()
def reject_revision(index: int) -> bool:
    """Reject revision by 1-based index in the Revisions collection."""
    return session.tool_reject_revision(index)


@server.tool()
def goto(target: Any) -> Dict[str, int]:
    """Go to a section (int), a bookmark (str), or text (str)."""
    return session.tool_goto(target)


@server.tool()
def save_as(path: str) -> str:
    """Save the active document to the given path."""
    return session.tool_save_as(path)


@server.tool()
def close(save: bool = False) -> bool:
    """Close the active document and quit Word."""
    return session.tool_close(save)


if __name__ == "__main__":
    # Run the MCP server over stdio
    server.run()

