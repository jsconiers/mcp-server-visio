# Local setup — READ FIRST

**This server cannot run on this Mac.** It is staged here, not installed.

`pywin32` publishes wheels for `win32`, `win_amd64`, `win_arm64` only — there is no
macOS build, and there is no sdist that would build one. It drives Visio through
Windows COM, and Visio has no macOS desktop app. `import win32com.client` fails here
with `ModuleNotFoundError`.

Deliberately **not** registered in `claude_desktop_config.json` — it would fail on
every Claude Desktop launch and show a permanently broken server.

## To actually run it (Windows box / Parallels VM / ICE workstation)

Requires: Windows, Microsoft Visio **Professional** (licensed, launched once so COM
registers), Python 3.11+, and the Azure stencils from
https://learn.microsoft.com/en-us/azure/architecture/icons/ extracted into `My Shapes`.

```powershell
git clone -b feat/open-diagram-lifecycle-instructions https://github.com/jsconiers/mcp-server-visio.git
cd mcp-server-visio
pip install -r requirements.txt
```

Use a real CPython install from python.org, **not** a uv-managed interpreter —
pywin32's post-install COM registration step is unreliable under uv on Windows.

Then in `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "visio": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\path\\to\\mcp-server-visio\\server.py"],
      "env": {}
    }
  }
}
```

## Status of the four patches

Applied on branch `feat/open-diagram-lifecycle-instructions`; syntax-checked only.
**None of it has been exercised against live Visio COM** — there is no machine here
that can. Test in this order, since each gates the next:

1. `open_diagram` on an existing .vsdx → then `list_shapes` returns its shapes
2. Open a .vsdx by hand in Visio first, then start the server → confirm it *attaches*
   rather than spawning a second instance
3. Ctrl-C the server with that hand-opened document unsaved → **your Visio must survive**
4. Let the server spawn its own Visio, then Ctrl-C → that instance should exit cleanly,
   no orphaned VISIO.EXE in Task Manager, no hung "Save changes?" dialog

Step 3 is the one that matters. If it fails, `_owns_app` isn't being set correctly and
it will eat unsaved work.
