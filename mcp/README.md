Word COM MCP Server
===================

This folder contains a Model Context Protocol (MCP) server that exposes native Microsoft Word COM automation on Windows.

Tools (namespace: `word`)
- `word.open(path)`: Open a .docx file (copy découpée) for editing.
- `word.set_track_changes(on=True)`: Enable/disable tracked changes.
- `word.insert_comment(anchor, text)`: Insert a Word comment at the given anchor.
- `word.write_revision(anchor, new_text)`: Replace/insert text at anchor (tracked).
- `word.accept_revision(index)`: Accept a revision by 1-based index.
- `word.reject_revision(index)`: Reject a revision by 1-based index.
- `word.goto(target)`: Go to a section (int), bookmark (str), or text (str).
- `word.save_as(path)`: Save the active document to path (must not overwrite original).
- `word.close(save=False)`: Close the document and quit Word.

Anchors
- `{ "find": "text" }` → first match in document
- `{ "bookmark": "Name" }` → existing bookmark
- `{ "range": [start, end] }` → Word character offsets (1-based)

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
- Safety: the server refuses `save_as` to the same path that was opened.
- Always work on the découpage copy provided by the app; do not open the original in write mode.

