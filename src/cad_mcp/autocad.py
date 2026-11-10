"""Small AutoCAD ActiveX adapter; imported lazily so tools can be inspected off-host."""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any

from .errors import CadError


def _com():
    try:
        import win32com.client
        return win32com.client
    except ImportError as exc:
        raise CadError("pywin32 is required to connect to AutoCAD on Windows") from exc


class AutoCAD:
    def __init__(self) -> None:
        self.app: Any = None
        self.doc: Any = None

    def connect(self) -> None:
        com = _com()
        try:
            if self.app is None:
                self.app = com.GetActiveObject("AutoCAD.Application")
            self.doc = self.app.ActiveDocument
        except Exception as exc:
            raise CadError("No active AutoCAD session was found. Open AutoCAD and a drawing first.") from exc

    def modelspace(self):
        self.connect()
        return self.doc.ModelSpace

    def layout_spaces(self):
        self.connect()
        return [layout.Block for layout in self.doc.Layouts]

    def entity(self, handle: str):
        self.connect()
        try:
            return self.doc.HandleToObject(str(handle))
        except Exception as exc:
            raise CadError(f"No entity found for handle {handle!r}") from exc

    @staticmethod
    def double_array(values: Any):
        import pythoncom
        from win32com.client import VARIANT
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in values))

    @staticmethod
    def object_array(values: Any):
        import pythoncom
        from win32com.client import VARIANT
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, list(values))

    @staticmethod
    def point(value: list[float]):
        if len(value) not in (2, 3):
            raise CadError("A point must contain 2 or 3 coordinates")
        xyz = tuple(float(v) for v in value)
        return (xyz[0], xyz[1], xyz[2] if len(xyz) == 3 else 0.0)

    def add_line(self, start: list[float], end: list[float]) -> str:
        a, b = self.point(start), self.point(end)
        obj = self.modelspace().AddLine(self.double_array(a), self.double_array(b))
        return str(obj.Handle)

    def add_circle(self, center: list[float], radius: float) -> str:
        if radius <= 0:
            raise CadError("Radius must be greater than zero")
        obj = self.modelspace().AddCircle(self.double_array(self.point(center)), float(radius))
        return str(obj.Handle)

    def add_rectangle(self, first: list[float], opposite: list[float]) -> str:
        x1, y1, z1 = self.point(first)
        x2, y2, z2 = self.point(opposite)
        if x1 == x2 or y1 == y2:
            raise CadError("Rectangle corners must differ in both X and Y")
        # A closed lightweight polyline is a native, editable rectangle.
        points = ((x1, y1), (x2, y1), (x2, y2), (x1, y2))
        flat = tuple(c for point in points for c in point)
        obj = self.modelspace().AddLightWeightPolyline(self.double_array(flat))
        obj.Closed = True
        obj.Elevation = (z1 + z2) / 2
        return str(obj.Handle)

    def move(self, handle: str, displacement: list[float]) -> None:
        obj = self.entity(handle)
        dx, dy, dz = self.point(displacement)
        obj.Move(self.double_array((0.0, 0.0, 0.0)), self.double_array((dx, dy, dz)))

    def rotate(self, handle: str, base: list[float], angle_degrees: float) -> None:
        obj = self.entity(handle)
        obj.Rotate(self.double_array(self.point(base)), float(angle_degrees) * 3.141592653589793 / 180.0)

    def erase(self, handle: str) -> None:
        self.entity(handle).Delete()

    @staticmethod
    def _get(obj: Any, name: str, default: Any = None) -> Any:
        try:
            value = getattr(obj, name)
            if isinstance(value, (str, int, float, bool)) or value is None:
                return value
            if isinstance(value, (tuple, list)):
                return [AutoCAD._json_value(item) for item in value]
            return str(value)
        except Exception:
            return default

    @staticmethod
    def _json_value(value: Any) -> Any:
        if isinstance(value, (str, int, float, bool)) or value is None:
            return value
        if isinstance(value, (tuple, list)):
            return [AutoCAD._json_value(item) for item in value]
        return str(value)

    def drawing_info(self) -> dict[str, Any]:
        self.connect()
        doc = self.doc
        result = {
            "name": self._get(doc, "Name"),
            "full_path": self._get(doc, "FullName"),
            "active_space": self._get(doc, "ActiveSpace"),
            "model_space_entity_count": self._get(doc.ModelSpace, "Count", 0),
            "paper_space_entity_count": self._get(doc.PaperSpace, "Count", 0),
        }
        try:
            summary = doc.SummaryInfo
            for key, prop in (("title", "Title"), ("subject", "Subject"),
                              ("author", "Author"), ("keywords", "Keywords"),
                              ("comments", "Comments"), ("last_saved_by", "LastSavedBy")):
                result[key] = self._get(summary, prop, "")
        except Exception:
            pass
        return result

    def list_layers(self) -> list[dict[str, Any]]:
        self.connect()
        rows = []
        for layer in self.doc.Layers:
            rows.append({
                "name": self._get(layer, "Name"), "color_index": self._get(layer, "Color"),
                "on": self._get(layer, "LayerOn"), "frozen": self._get(layer, "Freeze"),
                "locked": self._get(layer, "Lock"), "linetype": self._get(layer, "Linetype"),
                "lineweight": self._get(layer, "Lineweight"),
                "plot": self._get(layer, "Plottable"),
            })
        return rows

    def list_blocks(self, limit: int = 200, include_anonymous: bool = False) -> list[dict[str, Any]]:
        self.connect()
        rows = []
        for block in self.doc.Blocks:
            name = self._get(block, "Name", "")
            if not include_anonymous and str(name).startswith("*U"):
                continue
            rows.append({
                "name": name, "entity_count": self._get(block, "Count", 0),
                "is_xref": self._get(block, "IsXRef", False),
                "is_layout": self._get(block, "IsLayout", False),
                "path": self._get(block, "Path", ""),
                "origin": self._get(block, "Origin"),
            })
            if len(rows) >= max(1, min(int(limit), 1000)):
                break
        return rows

    def list_xrefs(self) -> list[dict[str, Any]]:
        self.connect()
        rows = []
        for block in self.doc.Blocks:
            if self._get(block, "IsXRef", False):
                rows.append({
                    "name": self._get(block, "Name"), "path": self._get(block, "Path", ""),
                    "is_unloaded": self._get(block, "IsUnloaded", None),
                    "is_resolved": self._get(block, "IsResolved", None),
                    "is_overlay": self._get(block, "IsOverlay", None),
                })
        return rows

    def _entity_data(self, obj: Any) -> dict[str, Any]:
        fields = ("ObjectName", "Handle", "Layer", "Linetype", "Color", "Lineweight",
                  "Visible", "Length", "Area", "Radius", "TextString", "Name", "EffectiveName",
                  "InsertionPoint", "StartPoint", "EndPoint", "Center", "Radius", "Rotation",
                  "ScaleFactor", "XScaleFactor", "YScaleFactor", "ZScaleFactor", "Closed")
        data = {key: self._get(obj, key) for key in fields}
        return {key: value for key, value in data.items() if value is not None}

    def get_entity(self, handle: str) -> dict[str, Any]:
        return self._entity_data(self.entity(handle))

    def list_entities(self, limit: int = 100, layer: str = "", type_name: str = "", offset: int = 0) -> list[dict[str, Any]]:
        self.connect()
        max_rows, offset = max(1, min(int(limit), 1000)), max(0, int(offset))
        rows, skipped = [], 0
        for obj in self.doc.ModelSpace:
            if layer and str(self._get(obj, "Layer", "")).casefold() != layer.casefold():
                continue
            if type_name and str(self._get(obj, "ObjectName", "")).casefold() != type_name.casefold():
                continue
            if skipped < offset:
                skipped += 1
                continue
            rows.append(self._entity_data(obj))
            if len(rows) >= max_rows:
                break
        return rows

    def create_block(self, name: str, handles: list[str], base: list[float]) -> None:
        if not name.strip():
            raise CadError("Block name cannot be empty")
        self.connect()
        try:
            self.doc.Blocks.Item(name)
        except Exception:
            block = self.doc.Blocks.Add(self.double_array(self.point(base)), name)
        else:
            raise CadError(f"Block {name!r} already exists")
        for handle in handles:
            obj = self.entity(handle)
            # CopyObjects is the supported COM API for moving entities into a block.
            self.doc.CopyObjects(self.object_array([obj]), block)
            obj.Delete()

    def insert_block(self, name: str, insertion: list[float], scale: float = 1.0, rotation_degrees: float = 0.0) -> str:
        if scale <= 0:
            raise CadError("Scale must be greater than zero")
        self.connect()
        try:
            block = self.doc.Blocks.Item(name)
            if block.IsLayout:
                raise CadError(f"Block {name!r} does not exist")
        except Exception as exc:
            raise CadError(f"Block {name!r} does not exist") from exc
        obj = self.modelspace().InsertBlock(
            self.double_array(self.point(insertion)), name, float(scale), float(scale), float(scale),
            float(rotation_degrees) * 3.141592653589793 / 180.0,
        )
        return str(obj.Handle)

    def run_lisp(self, path: str) -> None:
        file = Path(path).expanduser().resolve()
        if file.suffix.lower() not in {".lsp", ".fas", ".vlx"} or not file.is_file():
            raise CadError("LISP add-on must be an existing .lsp, .fas, or .vlx file")
        self.connect()
        self.doc.SendCommand(f'(load "{str(file).replace(chr(92), "/")}")\n')

    def load_addon(self, path: str) -> None:
        file = Path(path).expanduser().resolve()
        suffix = file.suffix.lower()
        commands = {".dll": "_.NETLOAD", ".arx": "_.ARXLOAD", ".dbx": "_.ARXLOAD"}
        if suffix not in commands or not file.is_file():
            raise CadError("Add-on must be an existing .dll (.NET) or .arx/.dbx (ObjectARX) file")
        self.connect()
        clean = str(file).replace(chr(92), "/")
        self.doc.SendCommand(f'{commands[suffix]} "{clean}"\n')

    def refedit_start(self, handle: str) -> None:
        """Start REFEDIT for a selected block reference via a temporary selection set."""
        self.connect()
        obj = self.entity(handle)
        if str(obj.ObjectName).lower() not in {"acdbblockreference", "acdbminsertblock"}:
            raise CadError(f"Entity {handle!r} is not a block reference")
        # AutoLISP resolves the handle to an entity name; command-line selection
        # does not accept a raw handle as an object pick.
        self.doc.SendCommand(
            f'(command "_.REFEDIT" (ssname (ssget "_X" '
            f"'((5 . \"{obj.Handle}\"))) 0) \"\")\n"
        )

    def refedit_save(self) -> None:
        self.connect()
        self.doc.SendCommand("_.REFCLOSE\n_Save\n")

    def refedit_discard(self) -> None:
        self.connect()
        self.doc.SendCommand("_.REFCLOSE\n_Discard\n")

    def run_command(self, command_name: str) -> None:
        """Invoke a named command registered by a loaded add-on."""
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.-]{0,127}", command_name):
            raise CadError("Command name contains unsupported characters")
        self.connect()
        self.doc.SendCommand(f"{command_name}\n")

    def copy_entity(self, handle: str, displacement: list[float]) -> str:
        source = self.entity(handle)
        duplicate = source.Copy()
        dx, dy, dz = self.point(displacement)
        duplicate.Move(self.double_array((0.0, 0.0, 0.0)), self.double_array((dx, dy, dz)))
        return str(duplicate.Handle)

    def mirror_entity(self, handle: str, first: list[float], second: list[float]) -> str:
        mirrored = self.entity(handle).Mirror(self.double_array(self.point(first)), self.double_array(self.point(second)))
        return str(mirrored.Handle)

    def scale_entity(self, handle: str, base: list[float], factor: float) -> None:
        if factor <= 0:
            raise CadError("Scale factor must be greater than zero")
        self.entity(handle).ScaleEntity(self.double_array(self.point(base)), float(factor))

    def find_entities(self, limit: int = 100, offset: int = 0, layer: str = "",
                      type_name: str = "", block_name: str = "") -> dict[str, Any]:
        self.connect()
        limit, offset = max(1, min(int(limit), 1000)), max(0, int(offset))
        rows, total = [], 0
        for space in self.layout_spaces():
            for obj in space:
                data = self._entity_data(obj)
                if layer and str(data.get("Layer", "")).casefold() != layer.casefold():
                    continue
                if type_name and str(data.get("ObjectName", "")).casefold() != type_name.casefold():
                    continue
                if block_name and str(data.get("EffectiveName", data.get("Name", ""))).casefold() != block_name.casefold():
                    continue
                if total >= offset and len(rows) < limit:
                    rows.append(data)
                total += 1
        return {"offset": offset, "returned": len(rows), "total_matches": total,
                "has_more": offset + len(rows) < total, "entities": rows}

    def set_entity_properties(self, handle: str, properties: dict[str, Any]) -> dict[str, Any]:
        allowed = {"Layer", "Color", "Linetype", "Lineweight", "Visible", "TextString"}
        if not properties or set(properties) - allowed:
            raise CadError(f"Provide supported properties only: {', '.join(sorted(allowed))}")
        self.connect()
        obj = self.entity(handle)
        for key, value in properties.items():
            if key == "Layer":
                try:
                    self.doc.Layers.Item(str(value))
                except Exception as exc:
                    raise CadError(f"Layer {value!r} does not exist") from exc
            setattr(obj, key, value)
        return self._entity_data(obj)

    def list_layouts(self) -> list[dict[str, Any]]:
        self.connect()
        result = []
        for layout in self.doc.Layouts:
            result.append({"name": self._get(layout, "Name"),
                           "model_type": self._get(layout, "ModelType"),
                           "tab_order": self._get(layout, "TabOrder"),
                           "entity_count": self._get(layout.Block, "Count", 0)})
        return result

    def list_block_references(self, limit: int = 200, offset: int = 0) -> dict[str, Any]:
        self.connect()
        limit, offset = max(1, min(int(limit), 1000)), max(0, int(offset))
        rows = []
        for space in self.layout_spaces():
            for obj in space:
                if str(self._get(obj, "ObjectName", "")).lower() != "acdbblockreference":
                    continue
                data = self._entity_data(obj)
                data["attributes"] = []
                try:
                    if obj.HasAttributes:
                        data["attributes"] = [{"tag": self._get(att, "TagString"),
                                               "value": self._get(att, "TextString")}
                                              for att in obj.GetAttributes()]
                except Exception:
                    pass
                rows.append(data)
        page = rows[offset:offset + limit]
        return {"offset": offset, "total": len(rows), "references": page,
                "has_more": offset + len(page) < len(rows)}

    def create_polyline(self, points: list[list[float]], closed: bool = False) -> str:
        if len(points) < (3 if closed else 2):
            raise CadError("Polyline needs at least two points, or three when closed")
        xyz = [tuple(self.point(p)) for p in points]
        if max(p[2] for p in xyz) - min(p[2] for p in xyz) > 1e-9:
            raise CadError("Lightweight polyline points must share one Z elevation")
        flat = tuple(c for p in xyz for c in p[:2])
        obj = self.modelspace().AddLightWeightPolyline(self.double_array(flat))
        obj.Closed = bool(closed)
        obj.Elevation = xyz[0][2]
        return str(obj.Handle)

    def create_arc(self, center: list[float], radius: float, start_degrees: float,
                   end_degrees: float) -> str:
        import math
        if radius <= 0:
            raise CadError("Radius must be greater than zero")
        obj = self.modelspace().AddArc(self.double_array(self.point(center)), float(radius),
                                       math.radians(start_degrees), math.radians(end_degrees))
        return str(obj.Handle)

    def create_text(self, text: str, insertion: list[float], height: float,
                    rotation_degrees: float = 0.0) -> str:
        import math
        if height <= 0:
            raise CadError("Text height must be greater than zero")
        obj = self.modelspace().AddText(str(text), self.double_array(self.point(insertion)), float(height))
        obj.Rotation = math.radians(rotation_degrees)
        return str(obj.Handle)

    def create_aligned_dimension(self, first: list[float], second: list[float],
                                 dimension_line_point: list[float]) -> str:
        obj = self.modelspace().AddDimAligned(self.double_array(self.point(first)),
                                              self.double_array(self.point(second)),
                                              self.double_array(self.point(dimension_line_point)))
        obj.Update()
        return str(obj.Handle)

    def create_hatch(self, boundary_handle: str, pattern: str = "SOLID",
                     scale: float = 1.0, angle_degrees: float = 0.0) -> str:
        import math
        if scale <= 0:
            raise CadError("Hatch scale must be greater than zero")
        boundary = self.entity(boundary_handle)
        if not self._get(boundary, "Closed", False):
            raise CadError("Hatch boundary must be a closed polyline")
        hatch = self.modelspace().AddHatch(0, pattern, True)
        import pythoncom
        from win32com.client import VARIANT
        outer_loop = self.object_array([boundary])
        hatch.AppendOuterLoop(outer_loop)
        hatch.PatternScale = float(scale)
        hatch.PatternAngle = math.radians(angle_degrees)
        hatch.Evaluate()
        return str(hatch.Handle)

    def create_layer(self, name: str, color_index: int = 7,
                     linetype: str = "Continuous") -> str:
        if not name.strip():
            raise CadError("Layer name cannot be empty")
        self.connect()
        try:
            layer = self.doc.Layers.Add(name)
            layer.Color = int(color_index)
            layer.Linetype = linetype
            return str(layer.Name)
        except Exception as exc:
            raise CadError(f"Could not create layer {name!r}: {exc}") from exc

    def set_layer_state(self, name: str, on: bool | None = None,
                        frozen: bool | None = None, locked: bool | None = None) -> dict[str, Any]:
        self.connect()
        try:
            layer = self.doc.Layers.Item(name)
        except Exception as exc:
            raise CadError(f"Layer {name!r} does not exist") from exc
        for prop, value in (("LayerOn", on), ("Freeze", frozen), ("Lock", locked)):
            if value is not None:
                setattr(layer, prop, bool(value))
        return {"name": layer.Name, "on": layer.LayerOn,
                "frozen": layer.Freeze, "locked": layer.Lock}

    def delete_layer(self, name: str) -> None:
        self.connect()
        if name.casefold() in {"0", "defpoints"}:
            raise CadError("Layer 0 and Defpoints cannot be deleted")
        try:
            self.doc.Layers.Item(name).Delete()
        except Exception as exc:
            raise CadError(f"Could not delete layer {name!r}; it may be current or contain entities") from exc

    def xref_action(self, name: str, action: str) -> None:
        self.connect()
        try:
            block = self.doc.Blocks.Item(name)
        except Exception as exc:
            raise CadError(f"Xref {name!r} does not exist") from exc
        if not self._get(block, "IsXRef", False):
            raise CadError(f"Block {name!r} is not an external reference")
        if action == "reload":
            block.Reload()
        elif action == "unload":
            block.Unload()
        else:
            raise CadError("Action must be reload or unload")

    def save_drawing(self, path: str = "") -> str:
        self.connect()
        if path:
            target = Path(path).expanduser().resolve()
            if target.exists():
                raise CadError(f"Save As target already exists: {target}")
            self.doc.SaveAs(str(target))
        else:
            if not str(self.doc.FullName):
                raise CadError("This drawing has no file path yet; provide a new path for Save As")
            self.doc.Save()
        return str(self.doc.FullName)

    def undo_mark(self, begin: bool) -> None:
        self.connect()
        if begin:
            self.doc.StartUndoMark()
        else:
            self.doc.EndUndoMark()

    def run_command(self, command_name: str) -> None:
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.-]{0,127}", command_name):
            raise CadError("Command name contains unsupported characters")
        self.connect()
        self.doc.SendCommand(f"{command_name}\n")
