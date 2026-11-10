# AutoCAD 2020 MCP server

A local, stdio-based MCP server for basic drawing creation and editing in an
already-running AutoCAD 2020 session on Windows. The first adapter uses AutoCAD's
ActiveX/COM API; the MCP tools are kept separate so another CAD adapter can be
added later.

## Setup

Use 64-bit Python on the same Windows machine as AutoCAD. In PowerShell, from
this folder:

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
cad-mcp
```

Open AutoCAD 2020 and a drawing before calling drawing tools. Configure your MCP
client to launch `cad-mcp` with stdio transport. On first use, approve any
AutoCAD prompts in its UI (for example, add-on security prompts and trusted
locations).

## Tools

- Session and drawing: `connect_autocad`, `get_drawing_info`, `save_drawing`
- Read: `list_layers`, `list_blocks`, `list_xrefs`, `list_layouts`, `list_block_references`, `get_entity`, `list_entities`, `find_entities`
- Create: `create_line`, `create_circle`, `create_rectangle`, `create_polyline`, `create_arc`, `create_text`, `create_aligned_dimension`, `create_hatch`, `create_block`, `insert_block`
- Edit: `move_entity`, `copy_entity`, `rotate_entity`, `mirror_entity`, `scale_entity`, `erase_entity`, `set_entity_properties`
- Layers and xrefs: `create_layer`, `set_layer_state`, `delete_layer`, `xref_action`
- Blocks: `refedit_block`, `refedit_close`
- Add-ons: `load_lisp`, `load_addon`, `run_addon_command`
- Multi-step edits: `undo_group` (call with `begin=true` before changes and `begin=false` after)

`load_lisp` accepts `.lsp`, `.fas`, and `.vlx`. `load_addon` accepts .NET
`.dll` assemblies and ObjectARX `.arx`/`.dbx` modules. A loaded add-on can register
commands invoked through `run_addon_command`. AutoCAD's COM `SendCommand` is
asynchronous and does not expose command output, so this tool reports submission
rather than claiming the add-on command succeeded.

`find_entities` searches model and paper-space layouts with pagination and filters
for layer, COM object type, or block name. `list_entities` reads model space with
offset pagination. Entity results contain common properties; fields vary by entity
type and AutoCAD availability. Xref status fields are included only when exposed
by AutoCAD's COM object model.

## Reading drawings

Inspection tools read the active drawing through AutoCAD COM. The server reads DWG
content through the running AutoCAD session; it does not parse DWG files directly.

## Notes

- Geometry uses drawing units. Points can be `[x, y]` or `[x, y, z]`.
- Entity handles are returned by creation tools and can be passed to modify tools.
- REFEDIT is an interactive AutoCAD command. The server submits it to the active
  drawing; complete any prompts in the AutoCAD window, then call `refedit_close`.
- Keep add-ons in a trusted AutoCAD location and load only files you trust.
- The ActiveX bridge requires Windows, AutoCAD 2020 running in the interactive
  user session, and matching Python/AutoCAD COM bitness where applicable.
