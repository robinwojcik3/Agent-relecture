Word COM MCP Server
===================

This folder contains a Model Context Protocol (MCP) server that exposes native Microsoft Word COM automation on Windows.

Tools (namespace: `word`)
- `word.open_document(path)`: Open a COPY of the given `.docx` into `work/` and edit that copy (never the original).
- `word.enable_tracking(on=True)`: Enable/disable tracked changes on the active document.
- `word.add_comment(range, text, author?)`: Add a comment on `{start,end}` or first `{find:"pattern"}`; optional `author`.
- `word.replace_text_tracked(old_text, new_text, match_case=False, whole_word=False)`: Replace all occurrences with tracked changes; returns count.
- `word.insert_text_tracked(position, text)`: Insert text at 1-based character position with tracked changes; returns inserted range.
- `word.save_document()`: Save the working copy (refuses saving to the original path).
- `word.save_as(path)`: Save the active document to `output/…` (must not overwrite original).
- `word.close(discard=False)`: Close the document and quit Word.

Anchors
- `{ "find": "text" }` → first match in document
- `{ "range": [start, end] }` → Word character offsets (1-based)

Mapping (needs → implementation)
- Open découpe safely → `open_document()` (copies to `work/`, opens copy)
- Enable track changes → `enable_tracking()`
- Add Word comments → `add_comment()`
- Write tracked replacements → `replace_text_tracked()`
- Insert tracked text → `insert_text_tracked()`
- Save working doc → `save_document()`
- Save final deliverable → `save_as()` (use `output/…`)
- Close Word session → `close(discard)`

Install
- `pip install -r requirements.txt`
  - Requires: Python 3.10+, Windows, Microsoft Word, `pywin32`, `mcp`.

Run locally
- `python mcp/word_com_server.py`

Client configuration (examples)

ChatGPT (mcpServers JSON)
```
{
  "mcpServers": {
    "word-com": {
      "command": "python",
      "args": ["mcp/word_com_server.py"],
      "env": {}
    }
  }
}
```

Cursor / Windsurf (similar)
- Add a server named `word-com` pointing to `python mcp/word_com_server.py`.

Notes
- Single session: the server serializes access and allows only one active Word instance.
- Safety: refuses any write to the original (`save_document`/`save_as`).
- Results always directed to `work/` (working copy) and `output/` (final deliverable).
