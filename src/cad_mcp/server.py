"""MCP tool surface for the AutoCAD adapter."""

from __future__ import annotations

import logging
from typing import Any

from mcp.server.fastmcp import FastMCP

from .autocad import AutoCAD
from .errors import CadError

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("cad_mcp")
cad = AutoCAD()
mcp = FastMCP("AutoCAD 2020 MCP")


def _run(fn, *args):
    try:
        return fn(*args)
    except CadError as exc:
        return f"CAD error: {exc}"
    except Exception as exc:
        log.exception("AutoCAD operation failed")
        return f"AutoCAD operation failed: {exc}"


@mcp.tool()
def connect_autocad() -> str:
    """Connect to the active AutoCAD 2020 session and report the drawing name."""
    return _run(lambda: (cad.connect(), f"Connected to {cad.doc.Name}")[1])


@mcp.tool()
def get_drawing_info() -> dict:
    """Read active drawing name, path, summary fields, and model/paper space entity counts."""
    return _run(cad.drawing_info)


@mcp.tool()
def list_layers() -> list[dict]:
    """List drawing layers with on/frozen/locked state, color, linetype, and plotting settings."""
    return _run(cad.list_layers)


@mcp.tool()
def list_blocks(limit: int = 200, include_anonymous: bool = False) -> list[dict]:
    """Inspect the drawing block table, including definition sizes and xref/layout flags."""
    return _run(cad.list_blocks, limit, include_anonymous)


@mcp.tool()
def list_xrefs() -> list[dict]:
    """List external reference definitions, paths, and resolution/overlay status when available."""
    return _run(cad.list_xrefs)


@mcp.tool()
def get_entity(handle: str) -> dict:
    """Read common properties and geometry fields for a drawing entity by handle."""
    return _run(cad.get_entity, handle)


@mcp.tool()
def list_entities(limit: int = 100, layer: str = "", type_name: str = "", offset: int = 0) -> list[dict]:
    """Read a paginated model-space slice with common geometry/properties; optionally filter by layer or COM ObjectName."""
    return _run(cad.list_entities, limit, layer, type_name, offset)


@mcp.tool()
def create_line(start: list[float], end: list[float]) -> str:
    """Create a line between two [x,y] or [x,y,z] points; return its entity handle."""
    return _run(cad.add_line, start, end)


@mcp.tool()
def create_circle(center: list[float], radius: float) -> str:
    """Create a circle at center with a positive radius; return its entity handle."""
    return _run(cad.add_circle, center, radius)


@mcp.tool()
def create_rectangle(first_corner: list[float], opposite_corner: list[float]) -> str:
    """Create a closed, editable rectangular polyline from two opposite corners."""
    return _run(cad.add_rectangle, first_corner, opposite_corner)


@mcp.tool()
def move_entity(handle: str, displacement: list[float]) -> str:
    """Move an entity by [dx,dy] or [dx,dy,dz], addressed by AutoCAD handle."""
    result = _run(cad.move, handle, displacement)
    return result if isinstance(result, str) else f"Moved entity {handle}."


@mcp.tool()
def rotate_entity(handle: str, base_point: list[float], angle_degrees: float) -> str:
    """Rotate an entity around a [x,y] or [x,y,z] point by degrees."""
    result = _run(cad.rotate, handle, base_point, angle_degrees)
    return result if isinstance(result, str) else f"Rotated entity {handle}."


@mcp.tool()
def erase_entity(handle: str) -> str:
    """Erase the entity with this AutoCAD handle."""
    result = _run(cad.erase, handle)
    return result if isinstance(result, str) else f"Erased entity {handle}."


@mcp.tool()
def create_block(name: str, entity_handles: list[str], base_point: list[float]) -> str:
    """Create a named block from existing entities and replace them with the block definition."""
    result = _run(cad.create_block, name, entity_handles, base_point)
    return result if isinstance(result, str) else f"Created block {name!r}."


@mcp.tool()
def insert_block(name: str, insertion_point: list[float], scale: float = 1.0,
                 rotation_degrees: float = 0.0) -> str:
    """Insert an existing named block; return its reference handle."""
    return _run(cad.insert_block, name, insertion_point, scale, rotation_degrees)


@mcp.tool()
def refedit_block(handle: str) -> str:
    """Start AutoCAD REFEDIT for a block reference identified by its handle."""
    result = _run(cad.refedit_start, handle)
    return result if isinstance(result, str) else "REFEDIT started in AutoCAD; finish or close it there."


@mcp.tool()
def refedit_close(save: bool = True) -> str:
    """Close the current REFEDIT session, saving changes or discarding them."""
    fn = cad.refedit_save if save else cad.refedit_discard
    result = _run(fn)
    return result if isinstance(result, str) else ("REFEDIT changes saved." if save else "REFEDIT changes discarded.")


@mcp.tool()
def load_lisp(path: str) -> str:
    """Load an existing AutoLISP .lsp, .fas, or .vlx routine by absolute path."""
    result = _run(cad.run_lisp, path)
    return result if isinstance(result, str) else f"Submitted LISP load for {path}."


@mcp.tool()
def load_addon(path: str) -> str:
    """Load a compiled AutoCAD add-on: .dll with NETLOAD or .arx/.dbx with ARXLOAD."""
    result = _run(cad.load_addon, path)
    return result if isinstance(result, str) else f"Submitted add-on load for {path}."


@mcp.tool()
def run_addon_command(command_name: str) -> str:
    """Submit a named command registered by a loaded AutoLISP, .NET, or ARX add-on."""
    result = _run(cad.run_command, command_name)
    return result if isinstance(result, str) else f"Submitted {command_name!r} to AutoCAD. AutoCAD COM does not return command output or errors."

@mcp.tool()
def find_entities(limit: int = 100, offset: int = 0, layer: str = "", type_name: str = "", block_name: str = "") -> dict:
    """Search model and paper space by layer, COM entity type, or referenced block name; results are paginated."""
    return _run(cad.find_entities, limit, offset, layer, type_name, block_name)


@mcp.tool()
def set_entity_properties(handle: str, properties: dict[str, Any]) -> dict:
    """Update supported entity properties: Layer, Color, Linetype, Lineweight, Visible, or TextString."""
    return _run(cad.set_entity_properties, handle, properties)


@mcp.tool()
def list_layouts() -> list[dict]:
    """List drawing layouts, tab order, and paper-space entity counts."""
    return _run(cad.list_layouts)


@mcp.tool()
def list_block_references(limit: int = 200, offset: int = 0) -> dict:
    """List placed block references with handles and attribute tag/value pairs, paginated."""
    return _run(cad.list_block_references, limit, offset)


@mcp.tool()
def copy_entity(handle: str, displacement: list[float]) -> str:
    """Copy an entity by handle and move the copy by [dx,dy] or [dx,dy,dz]."""
    return _run(cad.copy_entity, handle, displacement)


@mcp.tool()
def mirror_entity(handle: str, first_axis_point: list[float], second_axis_point: list[float]) -> str:
    """Mirror an entity across an axis defined by two points; return the new handle."""
    return _run(cad.mirror_entity, handle, first_axis_point, second_axis_point)


@mcp.tool()
def scale_entity(handle: str, base_point: list[float], factor: float) -> str:
    """Scale an entity uniformly about a base point by a positive factor."""
    result = _run(cad.scale_entity, handle, base_point, factor)
    return result if isinstance(result, str) else f"Scaled entity {handle}."


@mcp.tool()
def create_polyline(points: list[list[float]], closed: bool = False) -> str:
    """Create a lightweight polyline from coplanar [x,y] or [x,y,z] vertices."""
    return _run(cad.create_polyline, points, closed)


@mcp.tool()
def create_arc(center: list[float], radius: float, start_degrees: float, end_degrees: float) -> str:
    """Create an arc using center, radius, and start/end angles in degrees."""
    return _run(cad.create_arc, center, radius, start_degrees, end_degrees)


@mcp.tool()
def create_text(text: str, insertion_point: list[float], height: float, rotation_degrees: float = 0.0) -> str:
    """Create single-line text at an insertion point."""
    return _run(cad.create_text, text, insertion_point, height, rotation_degrees)


@mcp.tool()
def create_aligned_dimension(first_point: list[float], second_point: list[float], dimension_line_point: list[float]) -> str:
    """Create an aligned dimension between two points."""
    return _run(cad.create_aligned_dimension, first_point, second_point, dimension_line_point)


@mcp.tool()
def create_hatch(boundary_handle: str, pattern: str = "SOLID", scale: float = 1.0, angle_degrees: float = 0.0) -> str:
    """Hatch a closed polyline boundary with a named AutoCAD pattern."""
    return _run(cad.create_hatch, boundary_handle, pattern, scale, angle_degrees)


@mcp.tool()
def create_layer(name: str, color_index: int = 7, linetype: str = "Continuous") -> str:
    """Create a layer using an already loaded linetype."""
    return _run(cad.create_layer, name, color_index, linetype)


@mcp.tool()
def set_layer_state(name: str, on: bool | None = None, frozen: bool | None = None, locked: bool | None = None) -> dict:
    """Change a layer's on, frozen, and/or locked state."""
    return _run(cad.set_layer_state, name, on, frozen, locked)


@mcp.tool()
def delete_layer(name: str) -> str:
    """Delete an unused, non-current layer. AutoCAD rejects layers that cannot be removed."""
    result = _run(cad.delete_layer, name)
    return result if isinstance(result, str) else f"Deleted layer {name!r}."


@mcp.tool()
def xref_action(name: str, action: str) -> str:
    """Reload or unload a named external reference."""
    result = _run(cad.xref_action, name, action.lower())
    return result if isinstance(result, str) else f"Xref {name!r}: {action.lower()} submitted."


@mcp.tool()
def save_drawing(path: str = "") -> str:
    """Save the active drawing, or Save As to an explicit path when provided."""
    return _run(cad.save_drawing, path)


@mcp.tool()
def undo_group(begin: bool) -> str:
    """Begin or end an AutoCAD undo mark around a sequence of MCP edits."""
    result = _run(cad.undo_mark, begin)
    return result if isinstance(result, str) else ("Undo group started." if begin else "Undo group ended.")

def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
