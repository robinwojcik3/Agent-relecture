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
# Expose tools under the 'word' namespace
server = Server("word")


# ========== New normalized API tools ==========
@server.tool()
def open_document(path: str) -> str:
    """Open a COPY of the given .docx into work/ and edit that copy only."""
    # Create working copy in work/ and open it, never the original
    from pathlib import Path
    import shutil
    with session._lock:
        session._ensure_app()
        if session.doc is not None:
            session.doc.Close(SaveChanges=False)
            session.doc = None
        src = os.path.abspath(path)
        root = Path(__file__).resolve().parents[1]
        work_dir = root / "work"
        work_dir.mkdir(exist_ok=True)
        base = Path(src).name
        dst = work_dir / (Path(base).stem + "_MCP_WORK.docx")
        shutil.copy2(src, dst)
        session.original_path = src
        session.doc = session.app.Documents.Open(str(dst))
        return str(dst)


@server.tool()
def enable_tracking(on: bool = True) -> bool:
    """Enable/disable track changes on the active document."""
    with session._lock:
        if session.doc is None:
            raise RuntimeError("No document open")
        session.doc.TrackRevisions = bool(on)
        return bool(on)


@server.tool()
def add_comment(range: Dict[str, Any], text: str, author: Optional[str] = None) -> Dict[str, Any]:
    """Add a comment at a {start,end} range or first {find:"pattern"}."""
    with session._lock:
        if session.doc is None:
            raise RuntimeError("No document open")
        rng = session.doc.Content
        # Determine target range
        if isinstance(range, dict):
            if 'start' in range and 'end' in range:
                s, e = int(range['start']), int(range['end'])
                rng = session.doc.Range(Start=s, End=e)
            elif 'find' in range:
                text_find = str(range['find'] or '')
                f = rng.Find
                f.Text = text_find
                f.MatchCase = False
                f.MatchWholeWord = False
                if not f.Execute():
                    raise RuntimeError("Pattern not found for add_comment")
        # Temporarily set author if provided
        prev_name = None
        prev_init = None
        if author:
            prev_name = session.app.UserName
            prev_init = getattr(session.app, 'UserInitials', None)
            session.app.UserName = author
            try:
                initials = ''.join([w[0] for w in author.split() if w])[:3].upper()
                if prev_init is not None:
                    session.app.UserInitials = initials
            except Exception:
                pass
        try:
            c = session.doc.Comments.Add(rng, text)
            return {"index": int(c.Index), "start": int(rng.Start), "end": int(rng.End)}
        finally:
            if author:
                try:
                    session.app.UserName = prev_name
                    if prev_init is not None:
                        session.app.UserInitials = prev_init
                except Exception:
                    pass


@server.tool()
def replace_text_tracked(old_text: str, new_text: str, match_case: bool = False, whole_word: bool = False) -> int:
    """Replace all occurrences with tracked changes; returns count."""
    with session._lock:
        if session.doc is None:
            raise RuntimeError("No document open")
        count = 0
        content_end = session.doc.Content.End
        start = 0
        while True:
            rng = session.doc.Range(Start=start, End=content_end)
            f = rng.Find
            f.Text = old_text
            f.MatchCase = bool(match_case)
            f.MatchWholeWord = bool(whole_word)
            f.Wrap = 0  # wdFindStop
            if not f.Execute():
                break
            # Replace by setting Range.Text (tracked when TrackRevisions=True)
            rng.Text = new_text
            count += 1
            # Move start after the replaced range
            start = int(rng.End)
            content_end = session.doc.Content.End
        return count


@server.tool()
def insert_text_tracked(position: int, text: str) -> Dict[str, int]:
    """Insert text at 1-based character position; returns inserted range."""
    with session._lock:
        if session.doc is None:
            raise RuntimeError("No document open")
        pos = max(0, int(position))
        rng = session.doc.Range(Start=pos, End=pos)
        rng.Text = text
        return {"start": int(rng.Start), "end": int(rng.End)}


@server.tool()
def save_document() -> str:
    """Save the current working document (never the original)."""
    with session._lock:
        if session.doc is None:
            raise RuntimeError("No document open")
        current = os.path.abspath(session.doc.FullName)
        if session.original_path and os.path.abspath(current) == os.path.abspath(session.original_path):
            raise RuntimeError("Refusing to save to the original document")
        session.doc.Save()
        return current


# ========== Keep existing finalization tools ==========
@server.tool()
def save_as(path: str) -> str:
    """Save the active document to the given path; must not overwrite original."""
    with session._lock:
        if session.doc is None:
            raise RuntimeError("No document open")
        dest = os.path.abspath(path)
        if session.original_path and os.path.abspath(dest) == os.path.abspath(session.original_path):
            raise RuntimeError("Refusing to overwrite original document")
        # Ensure output directory exists
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        session.doc.SaveAs2(dest)
        return dest


@server.tool()
def close(discard: bool = False) -> bool:
    """Close the active document and quit Word. discard=True will not save pending changes."""
    with session._lock:
        if session.doc is not None:
            session.doc.Close(SaveChanges=not bool(discard))
            session.doc = None
        if session.app is not None:
            session.app.Quit()
            session.app = None
        pythoncom.CoUninitialize()
        return True


if __name__ == "__main__":
    # Run the MCP server over stdio
    server.run()
