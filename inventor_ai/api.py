"""
inventor_api.py — Autodesk Inventor 2026 COM automation wrapper.

All public methods accept dimensions in MILLIMETRES and convert to
Inventor's internal unit (CENTIMETRES) before any COM call.

Unit conversion rule:
    cm_value = mm_value / 10.0

The InventorAPI class is the single entry-point. Every method returns a
dict with at least {"success": bool, "error": str | None, ...}.

STL translator note:
    Inventor exposes STL export through an ApplicationAddIn whose
    DisplayName contains "STL". We locate it at runtime rather than
    hard-coding a CLSID, which avoids the DWF-translator CLSID confusion
    ({533E9A98-FC3B-11D4-8E7E-0010B541CD80} is the STL GUID; the user's
    brief mentioned ...CAA8 which is the DWF translator — never use that).

ProgID note:
    "Inventor.Application" (INV.APP) is the only registered ProgID; Autodesk does
    not expose version-suffixed ProgIDs (e.g. "Inventor.Application.26").
    The first installed / last-registered version is what you get.

makepy shortcut (run once to regenerate type stubs):
    python -m win32com.client.makepy "Autodesk Inventor Object Library"
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

# ---------------------------------------------------------------------------
# Raw integer fallbacks for Inventor constants.
# If makepy has been run these are unused; if not, we fall back to literals.
# ---------------------------------------------------------------------------
# Values verified at runtime from Inventor 2026.2 typelib
# (see gen_py typelib dict for confirmation).
_kPartDocumentObject = 12290            # ObjectTypeEnum.kPartDocumentObject
_kFileBrowseIOMechanism = 13059         # IOMechanismEnum.kFileBrowseIOMechanism

# Inventor origin WorkPlanes ordering (1-indexed, verified at runtime):
#   Item(1) = YZ Plane
#   Item(2) = XZ Plane
#   Item(3) = XY Plane
_PLANE_INDEX: dict[str, int] = {
    "YZ": 1,
    "XZ": 2,
    "XY": 3,
}

# PartFeatureExtentDirectionEnum
_EXTENT_DIR: dict[str, int] = {
    "positive":  20993,   # kPositiveExtentDirection
    "negative":  20994,   # kNegativeExtentDirection
    "symmetric": 20995,   # kSymmetricExtentDirection
}

# PartFeatureOperationEnum
_EXTENT_OP: dict[str, int] = {
    "join":      20481,   # kJoinOperation
    "cut":       20482,   # kCutOperation
    "intersect": 20483,   # kIntersectOperation
    "new_body":  20485,   # kNewBodyOperation
}

# STL resolution index for NameValueMap key "Resolution"
_STL_RESOLUTION: dict[str, int] = {
    "coarse": 3,   # Low
    "medium": 1,   # Medium
    "fine":   0,   # High
}

# Fallback CLSID for Inventor STL add-in (binary STL)
_STL_ADDIN_CLSID = "{533E9A98-FC3B-11D4-8E7E-0010B541CD80}"

# Fallback CLSID for Inventor STEP translator add-in (AP214)
_STEP_ADDIN_CLSID = "{90AF7F44-0C01-11D5-8E83-0010B541CD80}"

# AssemblyDocumentObject type for Documents.Add()
_kAssemblyDocumentObject = 12291

# Default Inventor 2026 metric assembly template
_DEFAULT_IAM_TEMPLATE = (
    r"C:\Users\Public\Documents\Autodesk\Inventor 2026\Templates\en-US\Metric\Standard (mm).iam"
)


class InventorAPI:
    """
    Clean wrapper around the Autodesk Inventor COM API.

    All length parameters are in millimetres at this surface; the wrapper
    converts to centimetres (Inventor's internal unit) before COM calls.
    """

    def __init__(self) -> None:
        self._app: Any = None
        # Maps part_name -> {"doc": PartDocument, "sketch": Sketch2D | None}
        self._docs: dict[str, dict[str, Any]] = {}

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _get_app(self) -> Any:
        """
        Return an active Inventor Application object.

        Strategy:
            1. If self._app is already set and alive, return it.
            2. Try GetActiveObject("Inventor.Application") — connects to a
               running instance without launching a new one.
            3. Fall back to Dispatch("Inventor.Application") — auto-launches
               if Inventor is registered on the system.
            4. Raise RuntimeError with a user-friendly message if both fail.
        """
        # Quick cache check. Inventor.Application has no .Version attribute
        # (that's an Excel-ism) — use SoftwareVersion.DisplayName which is
        # the documented path on the Inventor COM Application object.
        if self._app is not None:
            try:
                _ = self._app.SoftwareVersion.DisplayName  # liveness probe
                return self._app
            except Exception:
                self._app = None

        # EnsureDispatch builds/loads the typelib cache (required for
        # CastTo(PartDocument) etc.). It attaches to an existing instance
        # if one is running, else launches one.
        try:
            self._app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
            self._app.Visible = True
            _ = self._app.SoftwareVersion.DisplayName  # verify alive
            return self._app
        except Exception:
            self._app = None

        # Fallback: plain Dispatch (late-bound). CastTo will not be usable
        # in this mode but basic calls still work.
        try:
            candidate = win32com.client.Dispatch("Inventor.Application")
            candidate.Visible = True
            _ = candidate.SoftwareVersion.DisplayName
            self._app = candidate
            return self._app
        except Exception:
            self._app = None

        raise RuntimeError(
            "Inventor not available — launch Autodesk Inventor 2026 manually "
            "and retry. If Inventor is installed but not COM-registered, run as "
            "Administrator: "
            "\"C:\\Program Files\\Autodesk\\Inventor 2026\\Inventor.exe\" /regserver"
        )

    @staticmethod
    def _mm_to_cm(value_mm: float) -> float:
        """Convert millimetres to centimetres (Inventor's internal unit)."""
        return value_mm / 10.0

    def _get_doc_entry(self, part_name: str) -> dict[str, Any]:
        """Return the internal doc dict for part_name, raising KeyError if absent."""
        if part_name not in self._docs:
            raise KeyError(
                f"Part '{part_name}' not found. Call new_part() first."
            )
        return self._docs[part_name]

    def _wrap(self, method_name: str, **kwargs: Any) -> dict[str, Any]:
        """
        Decorator-equivalent: wraps a result dict, ensuring success/error keys.
        Not used externally — each public method has its own try/except.
        """
        return {"success": True, "error": None, **kwargs}

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def connect(self) -> dict[str, Any]:
        """
        Get or launch Inventor and return connection status.

        Returns:
            {"success": bool, "error": str|None, "version": str, "status": str}
        """
        try:
            app = self._get_app()
            try:
                version_str = app.SoftwareVersion.DisplayName
            except Exception:
                version_str = "unknown"
            return {
                "success": True,
                "error": None,
                "version": version_str,
                "status": "connected",
                "caption": app.Caption,
            }
        except Exception as e:
            return {"success": False, "error": str(e), "version": None, "status": "unavailable"}

    # ------------------------------------------------------------------
    # Document management
    # ------------------------------------------------------------------

    def new_part(self, part_name: str) -> dict[str, Any]:
        """
        Create a new PartDocument and register it under part_name.

        Args:
            part_name: Logical name used as a key throughout this session.

        Returns:
            {"success": bool, "error": str|None, "part_name": str}
        """
        try:
            app = self._get_app()
            # On Inventor 2026, Documents.Add + GetTemplateFile both raise
            # E_UNEXPECTED when a DesignProject with a non-default TemplatesPath
            # is active (reproducible with a .ipj project). The documented
            # work-around is to copy the template to a scratch path and open
            # it as an untitled working document.
            import os
            import shutil
            import tempfile
            import time

            template_candidates = [
                r"C:/Users/Public/Documents/Autodesk/Inventor 2026/Templates/en-US/Metric/Standard (mm).ipt",
                r"C:/Users/Public/Documents/Autodesk/Inventor 2026/Templates/en-US/Standard.ipt",
                r"C:/Users/Public/Documents/Autodesk/Inventor 2026/Templates/Standard.ipt",
            ]
            template_path = next((p for p in template_candidates if os.path.exists(p)), None)
            if template_path is None:
                raise FileNotFoundError(
                    "No Inventor Part template found. Checked: "
                    + "; ".join(template_candidates)
                )

            # Close any already-open documents whose DisplayName matches the
            # requested part_name — prevents "file already open" collisions
            # when re-running a workflow against the same logical part.
            i = app.Documents.Count
            while i >= 1:
                try:
                    d = app.Documents.Item(i)
                    if part_name.lower() in d.DisplayName.lower():
                        d.Close(True)  # SkipSave=True: no "save changes?" dialog
                except Exception:
                    pass
                i -= 1

            # Unique scratch path per call so repeat runs never collide.
            scratch_dir = Path(tempfile.gettempdir()) / "inventor_ai_scratch"
            scratch_dir.mkdir(parents=True, exist_ok=True)
            ts = time.strftime("%Y%m%d_%H%M%S")
            scratch_path = scratch_dir / f"{part_name}_{ts}.ipt"
            shutil.copyfile(template_path, scratch_path)

            doc_generic = app.Documents.Open(str(scratch_path), True)
            # Early-bound Documents.Open returns a base Document interface.
            # We need PartDocument to reach ComponentDefinition / Sketches / etc.
            try:
                doc = win32com.client.CastTo(doc_generic, "PartDocument")
            except Exception:
                doc = doc_generic  # late-bound fallback
            self._docs[part_name] = {
                "doc": doc,
                "sketch": None,
                "scratch_path": str(scratch_path),
                "template": template_path,
            }
            return {
                "success": True,
                "error": None,
                "part_name": part_name,
                "template": template_path,
                "scratch_path": str(scratch_path),
            }
        except Exception as e:
            return {"success": False, "error": str(e), "part_name": part_name}

    def list_open_documents(self) -> dict[str, Any]:
        """
        List all documents currently open in Inventor.

        Returns:
            {"success": bool, "error": str|None, "documents": [{"name": str, "path": str}]}
        """
        try:
            app = self._get_app()
            docs = []
            for i in range(1, app.Documents.Count + 1):
                d = app.Documents.Item(i)
                docs.append({"name": d.DisplayName, "path": d.FullFileName})
            return {"success": True, "error": None, "documents": docs}
        except Exception as e:
            return {"success": False, "error": str(e), "documents": []}

    def close_document(self, part_name: str, save: bool = False) -> dict[str, Any]:
        """
        Close a document registered under part_name.

        Args:
            part_name: Logical name of the part.
            save: Whether to save before closing.

        Returns:
            {"success": bool, "error": str|None}
        """
        try:
            entry = self._get_doc_entry(part_name)
            entry["doc"].Close(save)
            del self._docs[part_name]
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_part(self, part_name: str, file_path: str) -> dict[str, Any]:
        """
        Save a part document to a given path (.ipt).

        Args:
            part_name: Logical name of the part.
            file_path: Absolute path for the output .ipt file.

        Returns:
            {"success": bool, "error": str|None, "file_path": str}
        """
        try:
            entry = self._get_doc_entry(part_name)
            out = Path(file_path).resolve()
            out.parent.mkdir(parents=True, exist_ok=True)
            win_path = str(out).replace("/", "\\")
            # Ensure no other Inventor doc already holds this exact path;
            # close it if so, then remove the stale file so SaveAs doesn't
            # trip on a write conflict.
            app = self._get_app()
            i = app.Documents.Count
            while i >= 1:
                try:
                    d = app.Documents.Item(i)
                    if d.FullFileName and Path(d.FullFileName).resolve() == out:
                        if d is not entry["doc"]:
                            d.Close(True)  # SkipSave=True
                except Exception:
                    pass
                i -= 1
            if out.exists():
                try:
                    out.unlink()
                except Exception:
                    pass
            # SaveAs(FullFileName, SaveCopyAs=False) switches the doc to the
            # new path, which is what we want so subsequent calls operate
            # on the saved file instead of the scratch copy.
            entry["doc"].SaveAs(win_path, False)
            return {"success": True, "error": None, "file_path": win_path}
        except Exception as e:
            return {"success": False, "error": str(e), "file_path": file_path}

    # ------------------------------------------------------------------
    # Sketching
    # ------------------------------------------------------------------

    def new_sketch(self, part_name: str, plane: str = "XY") -> dict[str, Any]:
        """
        Create a new 2-D sketch on an origin work plane.

        Args:
            part_name: Logical name of the part.
            plane: One of "XY", "XZ", "YZ".

        Returns:
            {"success": bool, "error": str|None, "plane": str}
        """
        try:
            plane = plane.upper()
            if plane not in _PLANE_INDEX:
                raise ValueError(f"plane must be one of {list(_PLANE_INDEX)}; got '{plane}'")
            entry = self._get_doc_entry(part_name)
            comp_def = entry["doc"].ComponentDefinition
            # Origin work planes are 1-indexed in the API; our dict is 0-indexed
            work_plane = comp_def.WorkPlanes.Item(_PLANE_INDEX[plane])
            sketch = comp_def.Sketches.Add(work_plane)
            entry["sketch"] = sketch
            return {"success": True, "error": None, "plane": plane}
        except Exception as e:
            return {"success": False, "error": str(e), "plane": plane}

    def draw_circle(
        self,
        part_name: str,
        center_x_mm: float,
        center_y_mm: float,
        diameter_mm: float,
    ) -> dict[str, Any]:
        """
        Draw a circle on the active sketch.

        Args:
            part_name: Logical name of the part.
            center_x_mm: X coordinate of circle centre (mm).
            center_y_mm: Y coordinate of circle centre (mm).
            diameter_mm: Circle diameter (mm).

        Returns:
            {"success": bool, "error": str|None, "radius_mm": float}
        """
        try:
            entry = self._get_doc_entry(part_name)
            if entry["sketch"] is None:
                raise RuntimeError("No active sketch. Call new_sketch() first.")
            app = self._get_app()
            tg = app.TransientGeometry
            cx = self._mm_to_cm(center_x_mm)
            cy = self._mm_to_cm(center_y_mm)
            radius_cm = self._mm_to_cm(diameter_mm / 2.0)
            center_pt = tg.CreatePoint2d(cx, cy)
            entry["sketch"].SketchCircles.AddByCenterRadius(center_pt, radius_cm)
            return {"success": True, "error": None, "radius_mm": diameter_mm / 2.0}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def draw_line(
        self,
        part_name: str,
        x1_mm: float,
        y1_mm: float,
        x2_mm: float,
        y2_mm: float,
    ) -> dict[str, Any]:
        """
        Draw a line segment between two points on the active sketch.

        Args:
            part_name: Logical name of the part.
            x1_mm: X coordinate of line start (mm).
            y1_mm: Y coordinate of line start (mm).
            x2_mm: X coordinate of line end (mm).
            y2_mm: Y coordinate of line end (mm).

        Returns:
            {"success": bool, "error": str|None, "length_mm": float}
        """
        try:
            entry = self._get_doc_entry(part_name)
            if entry["sketch"] is None:
                raise RuntimeError("No active sketch. Call new_sketch() first.")
            app = self._get_app()
            tg = app.TransientGeometry
            x1 = self._mm_to_cm(x1_mm)
            y1 = self._mm_to_cm(y1_mm)
            x2 = self._mm_to_cm(x2_mm)
            y2 = self._mm_to_cm(y2_mm)
            p1 = tg.CreatePoint2d(x1, y1)
            p2 = tg.CreatePoint2d(x2, y2)
            entry["sketch"].SketchLines.AddByTwoPoints(p1, p2)
            # Calculate length in mm
            dx = x2_mm - x1_mm
            dy = y2_mm - y1_mm
            length_mm = (dx**2 + dy**2) ** 0.5
            return {"success": True, "error": None, "length_mm": length_mm}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def draw_rectangle(
        self,
        part_name: str,
        x_mm: float,
        y_mm: float,
        width_mm: float,
        height_mm: float,
    ) -> dict[str, Any]:
        """
        Draw an axis-aligned rectangle on the active sketch.

        Args:
            part_name: Logical name of the part.
            x_mm: X coordinate of lower-left corner (mm).
            y_mm: Y coordinate of lower-left corner (mm).
            width_mm: Width in mm.
            height_mm: Height in mm.

        Returns:
            {"success": bool, "error": str|None}
        """
        try:
            entry = self._get_doc_entry(part_name)
            if entry["sketch"] is None:
                raise RuntimeError("No active sketch. Call new_sketch() first.")
            app = self._get_app()
            tg = app.TransientGeometry
            x1 = self._mm_to_cm(x_mm)
            y1 = self._mm_to_cm(y_mm)
            x2 = self._mm_to_cm(x_mm + width_mm)
            y2 = self._mm_to_cm(y_mm + height_mm)
            p1 = tg.CreatePoint2d(x1, y1)
            p2 = tg.CreatePoint2d(x2, y2)
            entry["sketch"].SketchLines.AddAsTwoPointRectangle(p1, p2)
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Features
    # ------------------------------------------------------------------

    def extrude(
        self,
        part_name: str,
        distance_mm: float,
        direction: str = "positive",
        operation: str = "new_body",
    ) -> dict[str, Any]:
        """
        Extrude the first profile in the active sketch.

        Profile selection is naive: uses the first profile in the sketch.
        For complex sketches with multiple profiles, use the Inventor GUI.

        Args:
            part_name: Logical name of the part.
            distance_mm: Extrusion depth in mm.
            direction: "positive", "negative", or "symmetric".
            operation: "new_body", "join", "cut", or "intersect".

        Returns:
            {"success": bool, "error": str|None}
        """
        try:
            direction = direction.lower()
            operation = operation.lower()
            if direction not in _EXTENT_DIR:
                raise ValueError(f"direction must be one of {list(_EXTENT_DIR)}")
            if operation not in _EXTENT_OP:
                raise ValueError(f"operation must be one of {list(_EXTENT_OP)}")

            entry = self._get_doc_entry(part_name)
            if entry["sketch"] is None:
                raise RuntimeError("No active sketch. Call new_sketch() first.")

            sketch = entry["sketch"]
            profile = sketch.Profiles.AddForSolid()
            comp_def = entry["doc"].ComponentDefinition
            dist_cm = self._mm_to_cm(distance_mm)

            # Canonical Inventor extrude pattern (2018+):
            #   ext_def = ExtrudeFeatures.CreateExtrudeDefinition(profile, operation)
            #   ext_def.SetDistanceExtent(distance_cm, direction_enum)
            #   feat    = ExtrudeFeatures.Add(ext_def)
            extrudes = comp_def.Features.ExtrudeFeatures
            ext_def = extrudes.CreateExtrudeDefinition(profile, _EXTENT_OP[operation])
            ext_def.SetDistanceExtent(dist_cm, _EXTENT_DIR[direction])
            ext_feat = extrudes.Add(ext_def)
            return {"success": True, "error": None, "feature_name": ext_feat.Name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Parameters
    # ------------------------------------------------------------------

    def add_parameter(
        self,
        part_name: str,
        name: str,
        value: float,
        unit: str = "mm",
    ) -> dict[str, Any]:
        """
        Add a user parameter to the part.

        Args:
            part_name: Logical name of the part.
            name: Parameter name (Inventor identifier rules apply).
            value: Numeric value.
            unit: Unit string, e.g. "mm", "deg", "". Defaults to "mm".

        Returns:
            {"success": bool, "error": str|None, "name": str, "value": float}
        """
        try:
            entry = self._get_doc_entry(part_name)
            params = entry["doc"].ComponentDefinition.Parameters.UserParameters
            expression = f"{value} {unit}".strip() if unit else f"{value}"
            # AddByExpression(Name, Expression, Units) — Units can be a unit
            # string like "mm" or a UnitsTypeEnum int. Pass the string form
            # for correctness across unit systems.
            unit_arg = unit if unit else "ul"  # ul = unitless
            params.AddByExpression(name, expression, unit_arg)
            return {"success": True, "error": None, "name": name, "value": value}
        except Exception as e:
            return {"success": False, "error": str(e), "name": name, "value": value}

    def set_parameter(
        self,
        part_name: str,
        name: str,
        new_value: float,
    ) -> dict[str, Any]:
        """
        Update an existing user parameter's value.

        Args:
            part_name: Logical name of the part.
            name: Parameter name.
            new_value: New numeric value (same unit as when created).

        Returns:
            {"success": bool, "error": str|None, "name": str, "new_value": float}
        """
        try:
            entry = self._get_doc_entry(part_name)
            params = entry["doc"].ComponentDefinition.Parameters.UserParameters
            param = params.Item(name)
            param.Expression = str(new_value)
            return {"success": True, "error": None, "name": name, "new_value": new_value}
        except Exception as e:
            return {"success": False, "error": str(e), "name": name, "new_value": new_value}

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def export_stl(
        self,
        part_name: str,
        output_path: str,
        resolution: str = "medium",
    ) -> dict[str, Any]:
        """
        Export the part as an STL file using the Inventor STL translator add-in.

        Translator lookup strategy:
            1. Iterate ApplicationAddIns for DisplayName containing "STL"
               (case-insensitive) to avoid CLSID ambiguity.
            2. If not found, fall back to CLSID {533E9A98-FC3B-11D4-8E7E-0010B541CD80}.

        STL options set via NameValueMap:
            - OutputFileType = 0  (binary STL)
            - Resolution = 0|1|3  (High|Medium|Low)

        Note on units: Inventor's STL exporter works in the model's native
        units (cm). For Rocky DEM, import the STL and set unit to cm, OR
        use the ExportUnits option (value 4 = mm) if the add-in supports it.
        This wrapper sets ExportUnits=4 (mm) so the STL is in millimetres.

        Args:
            part_name: Logical name of the part.
            output_path: Absolute path for the .stl output file.
            resolution: "coarse", "medium", or "fine".

        Returns:
            {"success": bool, "error": str|None, "stl_path": str, "file_size_bytes": int}
        """
        try:
            resolution = resolution.lower()
            if resolution not in _STL_RESOLUTION:
                raise ValueError(f"resolution must be one of {list(_STL_RESOLUTION)}")

            entry = self._get_doc_entry(part_name)
            app = self._get_app()
            doc = entry["doc"]

            out = Path(output_path).resolve()
            out.parent.mkdir(parents=True, exist_ok=True)

            # --- Locate STL translator ---
            translator = None
            for i in range(1, app.ApplicationAddIns.Count + 1):
                addin = app.ApplicationAddIns.Item(i)
                try:
                    display_name = addin.DisplayName
                    if "stl" in display_name.lower():
                        translator = win32com.client.CastTo(addin, "TranslatorAddIn")
                        break
                except Exception:
                    continue

            # Fallback: use known CLSID
            if translator is None:
                try:
                    translator = win32com.client.CastTo(
                        app.ApplicationAddIns.ItemById(_STL_ADDIN_CLSID),
                        "TranslatorAddIn",
                    )
                except Exception:
                    pass

            if translator is None:
                raise RuntimeError(
                    "STL translator add-in not found. Ensure the STL add-in is "
                    "enabled in Inventor's Add-In Manager."
                )

            # --- Build translation objects ---
            context = app.TransientObjects.CreateTranslationContext()
            context.Type = _kFileBrowseIOMechanism  # 1

            options = app.TransientObjects.CreateNameValueMap()
            options.Add("OutputFileType", 0)          # 0 = binary STL
            options.Add("Resolution", _STL_RESOLUTION[resolution])
            options.Add("ExportUnits", 4)             # 4 = mm

            medium = app.TransientObjects.CreateDataMedium()
            # Inventor's STL translator requires a backslash-absolute path.
            medium.FileName = str(out).replace("/", "\\")

            # --- Export ---
            translator.SaveCopyAs(doc, context, options, medium)

            # Inventor's STL translator writes vertices in the model's
            # internal units (centimetres) regardless of ExportUnits=4.
            # Rescale the binary STL so the file is unambiguously in mm.
            import struct as _st
            with open(out, "rb") as _fh:
                _header = _fh.read(80)
                (_ntri,) = _st.unpack("<I", _fh.read(4))
                _body = _fh.read()
            _rescaled = bytearray()
            _rescaled.extend(_header)
            _rescaled.extend(_st.pack("<I", _ntri))
            _rec = 50  # 12*4 floats + 2-byte attr = 50 B/triangle
            for _i in range(_ntri):
                _tri = _body[_i * _rec:(_i + 1) * _rec]
                _floats = list(_st.unpack("<12f", _tri[:48]))
                # Floats 0-2 are the normal (unitless), floats 3-11 are 3
                # vertex XYZ triples — scale only the vertices.
                for _j in range(3, 12):
                    _floats[_j] *= 10.0
                _rescaled.extend(_st.pack("<12f", *_floats))
                _rescaled.extend(_tri[48:50])
            with open(out, "wb") as _fh:
                _fh.write(bytes(_rescaled))

            file_size = out.stat().st_size if out.exists() else -1
            return {
                "success": True,
                "error": None,
                "stl_path": str(out),
                "file_size_bytes": file_size,
            }
        except Exception as e:
            return {"success": False, "error": str(e), "stl_path": output_path}

    # ------------------------------------------------------------------
    # Mass properties
    # ------------------------------------------------------------------

    def get_mass_properties(self, part_name: str) -> dict[str, Any]:
        """
        Return volume, mass, surface area, and centre of mass for a part.

        Inventor MassProperties returns values in cm³ / cm² / kg.
        This method converts to mm³ / mm² but leaves mass in kg.

        Args:
            part_name: Logical name of the part.

        Returns:
            {
                "success": bool,
                "error": str|None,
                "volume_mm3": float,
                "mass_kg": float,
                "surface_area_mm2": float,
                "center_of_mass": {"x_mm": float, "y_mm": float, "z_mm": float},
            }
        """
        try:
            entry = self._get_doc_entry(part_name)
            mp = entry["doc"].ComponentDefinition.MassProperties
            # Inventor returns cm³ → ×1000 for mm³; cm² → ×100 for mm²
            vol_mm3 = mp.Volume * 1000.0
            area_mm2 = mp.Area * 100.0
            com = mp.CenterOfMass
            # CenterOfMass is in cm → convert each coordinate to mm
            return {
                "success": True,
                "error": None,
                "volume_mm3": vol_mm3,
                "mass_kg": mp.Mass,
                "surface_area_mm2": area_mm2,
                "center_of_mass": {
                    "x_mm": com.X * 10.0,
                    "y_mm": com.Y * 10.0,
                    "z_mm": com.Z * 10.0,
                },
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # High-level geometry builders
    # ------------------------------------------------------------------

    def create_cylinder(
        self,
        diameter_mm: float,
        height_mm: float,
        part_name: str,
        output_dir: str,
        axis: str = "Z",
    ) -> dict[str, Any]:
        """
        Create a solid cylinder, save as .ipt, and export as .stl.

        Workflow: new_part → sketch on plane perpendicular to axis → circle at
        origin → extrude along axis → save + STL.

        Args:
            diameter_mm: Outer diameter in mm.
            height_mm: Total height in mm.
            part_name: Logical name; also used as filename stem.
            output_dir: Directory for .ipt and .stl outputs.
            axis: Cylinder axis, one of "X", "Y", "Z". Default "Z".
                Selects sketch plane perpendicular to the axis
                (X→YZ, Y→XZ, Z→XY).

        Returns:
            {"success": bool, "error": str|None, "stl_path": str, "part_path": str, "volume_mm3": float}
        """
        try:
            axis = axis.upper()
            _SKETCH_PLANE_FOR_AXIS = {"X": "YZ", "Y": "XZ", "Z": "XY"}
            if axis not in _SKETCH_PLANE_FOR_AXIS:
                return {"success": False, "error": f"axis must be X, Y, or Z (got {axis!r})"}
            sketch_plane = _SKETCH_PLANE_FOR_AXIS[axis]

            out = Path(output_dir)
            out.mkdir(parents=True, exist_ok=True)
            part_path = str(out / f"{part_name}.ipt")
            stl_path = str(out / f"{part_name}.stl")

            for step, result in [
                ("new_part",      self.new_part(part_name)),
                ("new_sketch",    self.new_sketch(part_name, sketch_plane)),
                ("draw_circle",   self.draw_circle(part_name, 0.0, 0.0, diameter_mm)),
                ("extrude",       self.extrude(part_name, height_mm)),
                ("save_part",     self.save_part(part_name, part_path)),
            ]:
                if not result["success"]:
                    return {"success": False, "error": f"{step}: {result['error']}"}

            stl_result = self.export_stl(part_name, stl_path)
            if not stl_result["success"]:
                return {"success": False, "error": f"export_stl: {stl_result['error']}"}

            mp = self.get_mass_properties(part_name)
            volume = mp.get("volume_mm3", 0.0) if mp["success"] else 0.0

            return {
                "success": True,
                "error": None,
                "stl_path": stl_path,
                "part_path": part_path,
                "volume_mm3": volume,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_box(
        self,
        width_mm: float,
        depth_mm: float,
        height_mm: float,
        part_name: str,
        output_dir: str,
    ) -> dict[str, Any]:
        """
        Create a solid rectangular box, save as .ipt, and export as .stl.

        Rectangle is drawn with lower-left corner at origin on the XY plane.

        Args:
            width_mm: X dimension in mm.
            depth_mm: Y dimension in mm.
            height_mm: Z extrusion height in mm.
            part_name: Logical name; also used as filename stem.
            output_dir: Directory for outputs.

        Returns:
            {"success": bool, "error": str|None, "stl_path": str, "part_path": str, "volume_mm3": float}
        """
        try:
            out = Path(output_dir)
            out.mkdir(parents=True, exist_ok=True)
            part_path = str(out / f"{part_name}.ipt")
            stl_path = str(out / f"{part_name}.stl")

            for step, result in [
                ("new_part",       self.new_part(part_name)),
                ("new_sketch",     self.new_sketch(part_name, "XY")),
                ("draw_rectangle", self.draw_rectangle(part_name, 0.0, 0.0, width_mm, depth_mm)),
                ("extrude",        self.extrude(part_name, height_mm)),
                ("save_part",      self.save_part(part_name, part_path)),
            ]:
                if not result["success"]:
                    return {"success": False, "error": f"{step}: {result['error']}"}

            stl_result = self.export_stl(part_name, stl_path)
            if not stl_result["success"]:
                return {"success": False, "error": f"export_stl: {stl_result['error']}"}

            mp = self.get_mass_properties(part_name)
            volume = mp.get("volume_mm3", 0.0) if mp["success"] else 0.0

            return {
                "success": True,
                "error": None,
                "stl_path": stl_path,
                "part_path": part_path,
                "volume_mm3": volume,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_funnel(
        self,
        top_diameter_mm: float,
        bottom_diameter_mm: float,
        height_mm: float,
        part_name: str,
        output_dir: str,
    ) -> dict[str, Any]:
        """
        Create a conical funnel (truncated cone) by revolving a trapezoidal
        half-profile 360 degrees around the sketch Y-axis.

        Profile (drawn on XY plane, right-hand side only):
            P0 = (top_r, 0)        — top outer edge
            P1 = (top_r, 0)        — will extrude; actually forms the top face
            The half-profile is a 4-point closed polyline:
                bottom_r at z=0, top_r at z=height, then back along axis.

        Sketch layout (X = radial, Y = axial):
            A = (bottom_r,   0)
            B = (top_r,      height)
            C = (0,          height)
            D = (0,          0)
        The closed polyline A→B→C→D→A is revolved around the Y-axis (X=0 line).

        Args:
            top_diameter_mm: Inner/outer diameter at the wide top (mm).
            bottom_diameter_mm: Diameter at the narrow bottom outlet (mm).
            height_mm: Axial height of the funnel (mm).
            part_name: Logical name; also used as filename stem.
            output_dir: Directory for outputs.

        Returns:
            {"success": bool, "error": str|None, "stl_path": str, "part_path": str, "volume_mm3": float}
        """
        try:
            out = Path(output_dir)
            out.mkdir(parents=True, exist_ok=True)
            part_path = str(out / f"{part_name}.ipt")
            stl_path = str(out / f"{part_name}.stl")

            np_result = self.new_part(part_name)
            if not np_result["success"]:
                return {"success": False, "error": f"new_part: {np_result['error']}"}

            ns_result = self.new_sketch(part_name, "XY")
            if not ns_result["success"]:
                return {"success": False, "error": f"new_sketch: {ns_result['error']}"}

            entry = self._get_doc_entry(part_name)
            app = self._get_app()
            tg = app.TransientGeometry
            sketch = entry["sketch"]

            # Convert to cm
            bot_r = self._mm_to_cm(bottom_diameter_mm / 2.0)
            top_r = self._mm_to_cm(top_diameter_mm / 2.0)
            h = self._mm_to_cm(height_mm)

            # 4-point closed polyline (half-profile for revolution):
            # A=(bot_r, 0), B=(top_r, h), C=(0, h), D=(0, 0)
            lines = sketch.SketchLines
            pA = tg.CreatePoint2d(bot_r, 0.0)
            pB = tg.CreatePoint2d(top_r, h)
            pC = tg.CreatePoint2d(0.0,   h)
            pD = tg.CreatePoint2d(0.0,   0.0)

            lines.AddByTwoPoints(pA, pB)
            lines.AddByTwoPoints(pB, pC)
            lines.AddByTwoPoints(pC, pD)
            lines.AddByTwoPoints(pD, pA)

            # Revolution axis: line along the sketch Y-axis (X=0)
            axis_start = tg.CreatePoint2d(0.0, 0.0)
            axis_end   = tg.CreatePoint2d(0.0, h)
            axis_line  = lines.AddByTwoPoints(axis_start, axis_end)

            profile = sketch.Profiles.AddForSolid()
            comp_def = entry["doc"].ComponentDefinition

            # Full 360° revolve
            rev_feat = comp_def.Features.RevolveFeatures.AddFull(
                profile,
                axis_line,
                _EXTENT_OP["new_body"],
            )

            save_result = self.save_part(part_name, part_path)
            if not save_result["success"]:
                return {"success": False, "error": f"save_part: {save_result['error']}"}

            stl_result = self.export_stl(part_name, stl_path)
            if not stl_result["success"]:
                return {"success": False, "error": f"export_stl: {stl_result['error']}"}

            mp = self.get_mass_properties(part_name)
            volume = mp.get("volume_mm3", 0.0) if mp["success"] else 0.0

            return {
                "success": True,
                "error": None,
                "stl_path": stl_path,
                "part_path": part_path,
                "volume_mm3": volume,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def create_oedometer_container(
        self,
        outer_diameter_mm: float,
        inner_diameter_mm: float,
        height_mm: float,
        floor_thickness_mm: float,
        part_name: str,
        output_dir: str,
    ) -> dict[str, Any]:
        """Create a hollow cylindrical oedometer container by revolving a cup
        half-profile around the sketch Y-axis. Closed bottom, open top."""
        import math
        try:
            if inner_diameter_mm >= outer_diameter_mm:
                raise ValueError('inner_diameter_mm must be < outer_diameter_mm')
            if floor_thickness_mm <= 0 or floor_thickness_mm >= height_mm:
                raise ValueError('floor_thickness_mm must be in (0, height_mm)')
            out = Path(output_dir); out.mkdir(parents=True, exist_ok=True)
            part_path = str(out / f'{part_name}.ipt')
            stl_path = str(out / f'{part_name}.stl')
            np_result = self.new_part(part_name)
            if not np_result['success']:
                return {'success': False, 'error': f"new_part: {np_result['error']}"}
            ns_result = self.new_sketch(part_name, 'XY')
            if not ns_result['success']:
                return {'success': False, 'error': f"new_sketch: {ns_result['error']}"}
            entry = self._get_doc_entry(part_name)
            app = self._get_app()
            tg = app.TransientGeometry
            sketch = entry['sketch']
            outer_r = self._mm_to_cm(outer_diameter_mm / 2.0)
            inner_r = self._mm_to_cm(inner_diameter_mm / 2.0)
            h = self._mm_to_cm(height_mm)
            t = self._mm_to_cm(floor_thickness_mm)
            lines = sketch.SketchLines
            # Chain lines via .EndSketchPoint so Inventor shares endpoints and
            # recognises the polyline as a closed profile. Passing fresh
            # TransientGeometry points to every AddByTwoPoints creates
            # disconnected endpoints and Profiles.AddForSolid() fails.
            l1 = lines.AddByTwoPoints(tg.CreatePoint2d(0.0, 0.0),
                                      tg.CreatePoint2d(outer_r, 0.0))
            l2 = lines.AddByTwoPoints(l1.EndSketchPoint, tg.CreatePoint2d(outer_r, h))
            l3 = lines.AddByTwoPoints(l2.EndSketchPoint, tg.CreatePoint2d(inner_r, h))
            l4 = lines.AddByTwoPoints(l3.EndSketchPoint, tg.CreatePoint2d(inner_r, t))
            l5 = lines.AddByTwoPoints(l4.EndSketchPoint, tg.CreatePoint2d(0.0, t))
            lines.AddByTwoPoints(l5.EndSketchPoint, l1.StartSketchPoint)  # close
            profile = sketch.Profiles.AddForSolid()
            comp_def = entry['doc'].ComponentDefinition
            y_axis = comp_def.WorkAxes.Item(2)
            comp_def.Features.RevolveFeatures.AddFull(profile, y_axis, _EXTENT_OP['new_body'])
            save_result = self.save_part(part_name, part_path)
            if not save_result['success']:
                return {'success': False, 'error': f"save_part: {save_result['error']}"}
            stl_result = self.export_stl(part_name, stl_path)
            if not stl_result['success']:
                return {'success': False, 'error': f"export_stl: {stl_result['error']}"}
            mp = self.get_mass_properties(part_name)
            volume = mp.get('volume_mm3', 0.0) if mp['success'] else 0.0
            inner_vol = math.pi * (inner_diameter_mm / 2.0) ** 2 * (height_mm - floor_thickness_mm)
            return {'success': True, 'error': None, 'stl_path': stl_path, 'part_path': part_path,
                    'volume_mm3': volume, 'inner_volume_mm3': inner_vol}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    # ------------------------------------------------------------------
    # Diagnostics, escape hatch, and advanced features
    # ------------------------------------------------------------------

    def test_connection(self) -> dict[str, Any]:
        """Ping Inventor and report version + status. Useful as a smoke test."""
        try:
            app = self._get_app()
            return {
                "success": True,
                "error": None,
                "version": app.SoftwareVersion.DisplayName,
                "visible": app.Visible,
                "open_documents": app.Documents.Count,
                "registered_parts": list(self._docs.keys()),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def run_python(
        self,
        code: str,
        part_name: str | None = None,
    ) -> dict[str, Any]:
        """Execute arbitrary Python with Inventor COM objects pre-injected.

        Escape hatch for any Inventor operation we don't have a dedicated
        wrapper for. Modeled on AuraFriday Fusion-360-MCP-Server's
        execute_python operation.

        Pre-injected names in the code's namespace:
            app, doc, comp_def, sketch, tg, mm_to_cm,
            EXTENT_OP, EXTENT_DIR, PLANE_INDEX, api

        Set a variable named ``result`` in the code to return a value.
        Raw COM proxies are repr()'d (not returned) since they aren't
        JSON-serializable.
        """
        try:
            app = self._get_app()
            doc = None
            comp_def = None
            sketch = None
            if part_name is not None:
                entry = self._get_doc_entry(part_name)
                doc = entry["doc"]
                comp_def = doc.ComponentDefinition
                sketch = entry["sketch"]
            ns: dict[str, Any] = {
                "app": app,
                "doc": doc,
                "comp_def": comp_def,
                "sketch": sketch,
                "tg": app.TransientGeometry,
                "mm_to_cm": self._mm_to_cm,
                "EXTENT_OP": _EXTENT_OP,
                "EXTENT_DIR": _EXTENT_DIR,
                "PLANE_INDEX": _PLANE_INDEX,
                "api": self,
                "result": None,
            }
            exec(code, ns)  # noqa: S102 — intentional escape hatch
            result = ns.get("result")
            if result is not None and hasattr(result, "_oleobj_"):
                result = repr(result)
            return {"success": True, "error": None, "result": result}
        except Exception as e:
            return {"success": False, "error": f"{type(e).__name__}: {e}"}

    def fillet_all_edges(
        self,
        part_name: str,
        radius_mm: float,
    ) -> dict[str, Any]:
        """Round every edge of the part's first body with a uniform radius."""
        code = (
            "edge_set = app.TransientObjects.CreateEdgeCollection()\n"
            "body = comp_def.SurfaceBodies.Item(1)\n"
            "for edge in body.Edges:\n"
            "    edge_set.Add(edge)\n"
            'if edge_set.Count == 0:\n'
            '    raise RuntimeError("no edges on body")\n'
            "fillets = comp_def.Features.FilletFeatures\n"
            "fdef = fillets.CreateFilletDefinition()\n"
            f"fdef.AddConstantRadiusEdgeSet(edge_set, mm_to_cm({radius_mm}))\n"
            "feat = fillets.Add(fdef)\n"
            'result = {"feature_name": feat.Name, "edges_filleted": edge_set.Count}\n'
        )
        return self.run_python(code, part_name=part_name)

    def shell(
        self,
        part_name: str,
        thickness_mm: float,
        face_filter: str = "top",
    ) -> dict[str, Any]:
        """Hollow the part's first body, removing one face.

        Args:
            face_filter: which face to remove, one of:
                "top"    - highest Y centroid
                "bottom" - lowest Y centroid
                "+z"     - highest Z centroid
                "-z"     - lowest Z centroid

        Uses ThickenDirectionEnum.kInsideThickenDirection = 33793.
        """
        if face_filter not in ("top", "bottom", "+z", "-z"):
            return {
                "success": False,
                "error": f"face_filter must be top/bottom/+z/-z, got {face_filter!r}",
            }
        coord = 1 if face_filter in ("top", "bottom") else 2  # Y or Z (0-based)
        higher = face_filter in ("top", "+z")
        cmp_op = ">" if higher else "<"
        code = (
            "body = comp_def.SurfaceBodies.Item(1)\n"
            "chosen = None\n"
            "chosen_val = None\n"
            "for face in body.Faces:\n"
            "    box = face.RangeBox\n"
            f"    mid = (box.MinPoint.Coordinates[{coord}] + box.MaxPoint.Coordinates[{coord}]) / 2.0\n"
            f"    if chosen_val is None or (mid {cmp_op} chosen_val):\n"
            "        chosen = face\n"
            "        chosen_val = mid\n"
            "faces = app.TransientObjects.CreateObjectCollection()\n"
            "faces.Add(chosen)\n"
            "shells = comp_def.Features.ShellFeatures\n"
            f"feat = shells.Add(faces, mm_to_cm({thickness_mm}), 33793)\n"
            'result = {"feature_name": feat.Name, "face_centroid_cm": chosen_val}\n'
        )
        return self.run_python(code, part_name=part_name)

    def sweep(
        self,
        part_name: str,
        profile_sketch_idx: int,
        path_sketch_idx: int,
        operation: str = "new_body",
    ) -> dict[str, Any]:
        """Sweep a profile sketch along a path sketch (both 1-based indices)."""
        if operation not in _EXTENT_OP:
            return {"success": False, "error": f"operation must be one of {list(_EXTENT_OP)}"}
        op_val = _EXTENT_OP[operation]
        code = (
            f"prof_sk = comp_def.Sketches.Item({profile_sketch_idx})\n"
            f"path_sk = comp_def.Sketches.Item({path_sketch_idx})\n"
            "profile = prof_sk.Profiles.AddForSolid()\n"
            "path = path_sk.Path\n"
            "sweeps = comp_def.Features.SweepFeatures\n"
            f"sdef = sweeps.CreateSweepDefinition(0, profile, path, {op_val})\n"
            "feat = sweeps.Add(sdef)\n"
            'result = {"feature_name": feat.Name}\n'
        )
        return self.run_python(code, part_name=part_name)

    def mirror(
        self,
        part_name: str,
        plane: str = "XY",
    ) -> dict[str, Any]:
        """Mirror the most recently added solid feature across an origin plane."""
        if plane not in _PLANE_INDEX:
            return {"success": False, "error": f"plane must be one of {list(_PLANE_INDEX)}"}
        plane_idx = _PLANE_INDEX[plane]
        # PartFeatureExtentDirectionEnum.kIdenticalCompute = 20737
        code = (
            "feats = comp_def.Features\n"
            "if feats.Count == 0:\n"
            '    raise RuntimeError("no features to mirror")\n'
            "last_feat = feats.Item(feats.Count)\n"
            "feat_set = app.TransientObjects.CreateObjectCollection()\n"
            "feat_set.Add(last_feat)\n"
            f"mplane = comp_def.WorkPlanes.Item({plane_idx})\n"
            "mirrors = feats.MirrorFeatures\n"
            "feat = mirrors.Add(feat_set, mplane, 20737)\n"
            'result = {"feature_name": feat.Name, "mirrored": last_feat.Name}\n'
        )
        return self.run_python(code, part_name=part_name)

    def circular_pattern(
        self,
        part_name: str,
        count: int,
        axis: str = "Y",
        angle_deg: float = 360.0,
    ) -> dict[str, Any]:
        """Circular-pattern the most recent feature around an origin axis.

        Inventor WorkAxes Item index: 1=X, 2=Y, 3=Z (1-based).
        """
        axis_map = {"X": 1, "Y": 2, "Z": 3}
        if axis not in axis_map:
            return {"success": False, "error": "axis must be X, Y, or Z"}
        axis_idx = axis_map[axis]
        code = (
            "feats = comp_def.Features\n"
            "if feats.Count == 0:\n"
            '    raise RuntimeError("no features to pattern")\n'
            "last_feat = feats.Item(feats.Count)\n"
            "feat_set = app.TransientObjects.CreateObjectCollection()\n"
            "feat_set.Add(last_feat)\n"
            f"work_axis = comp_def.WorkAxes.Item({axis_idx})\n"
            "patterns = feats.CircularPatternFeatures\n"
            # CircularPatternFeatures.Add() in Inventor 2026's gencache stub
            # raises E_FAIL even with the documented 6-arg form. The
            # CreateDefinition + AddByDefinition path works reliably.
            f'cdef = patterns.CreateDefinition(feat_set, work_axis, True, "{count}", "{angle_deg} deg")\n'
            "feat = patterns.AddByDefinition(cdef)\n"
            'result = {"feature_name": feat.Name}\n'
        )
        return self.run_python(code, part_name=part_name)

    def rectangular_pattern(
        self,
        part_name: str,
        count_x: int,
        count_y: int,
        spacing_x_mm: float,
        spacing_y_mm: float,
        axis_x: str = "X",
        axis_y: str = "Z",
    ) -> dict[str, Any]:
        """Rectangular-pattern the most recent feature in two directions."""
        axis_map = {"X": 1, "Y": 2, "Z": 3}
        if axis_x not in axis_map or axis_y not in axis_map:
            return {"success": False, "error": "axis_x/axis_y must be X, Y, or Z"}
        ax_x = axis_map[axis_x]
        ax_y = axis_map[axis_y]
        # RectangularPatternFeatures.Add signature in Inventor 2026 (typelib):
        #   Add(ParentFeatures, XDirectionEntity, NaturalXDirection, XCount,
        #       XSpacing, XSpacingType=33537, XDirectionStartPoint=MISSING,
        #       YDirectionEntity, NaturalYDirection, YCount, YSpacing, ...)
        # XSpacingType 33537 = kDefaultPatternSpacing.
        # CreateDefinition only takes the X side (max 6 args) — direct Add()
        # is simpler.
        code = (
            "import pythoncom\n"
            "MISSING = pythoncom.Empty\n"
            "feats = comp_def.Features\n"
            "if feats.Count == 0:\n"
            '    raise RuntimeError("no features to pattern")\n'
            "last_feat = feats.Item(feats.Count)\n"
            "feat_set = app.TransientObjects.CreateObjectCollection()\n"
            "feat_set.Add(last_feat)\n"
            f"axx = comp_def.WorkAxes.Item({ax_x})\n"
            f"axy = comp_def.WorkAxes.Item({ax_y})\n"
            "feat = feats.RectangularPatternFeatures.Add(\n"
            "    feat_set, axx, True,\n"
            f'    "{count_x}", "{spacing_x_mm} mm",\n'
            "    33537, MISSING,\n"
            "    axy, True,\n"
            f'    "{count_y}", "{spacing_y_mm} mm",\n'
            ")\n"
            'result = {"feature_name": feat.Name}\n'
        )
        return self.run_python(code, part_name=part_name)

    def export_step(
        self,
        part_name: str,
        output_path: str,
    ) -> dict[str, Any]:
        """Export the part as a STEP AP214 file."""
        try:
            entry = self._get_doc_entry(part_name)
            app = self._get_app()
            doc = entry["doc"]
            out = Path(output_path).resolve()
            out.parent.mkdir(parents=True, exist_ok=True)

            translator = None
            for i in range(1, app.ApplicationAddIns.Count + 1):
                addin = app.ApplicationAddIns.Item(i)
                try:
                    if "step" in addin.DisplayName.lower():
                        translator = win32com.client.CastTo(addin, "TranslatorAddIn")
                        break
                except Exception:
                    continue
            if translator is None:
                try:
                    translator = win32com.client.CastTo(
                        app.ApplicationAddIns.ItemById(_STEP_ADDIN_CLSID),
                        "TranslatorAddIn",
                    )
                except Exception:
                    pass
            if translator is None:
                raise RuntimeError("STEP translator add-in not found")

            context = app.TransientObjects.CreateTranslationContext()
            context.Type = _kFileBrowseIOMechanism
            options = app.TransientObjects.CreateNameValueMap()
            options.Add("ApplicationProtocolType", 3)  # 3 = AP214 (auto draft)
            medium = app.TransientObjects.CreateDataMedium()
            medium.FileName = str(out).replace("/", "\\")
            translator.SaveCopyAs(doc, context, options, medium)
            size = out.stat().st_size if out.exists() else -1
            return {"success": True, "error": None, "step_path": str(out),
                    "file_size_bytes": size}
        except Exception as e:
            return {"success": False, "error": str(e), "step_path": output_path}

    def undo(self) -> dict[str, Any]:
        """Send a single Undo command to Inventor."""
        try:
            app = self._get_app()
            app.CommandManager.ControlDefinitions.Item("AppUndoCmd").Execute()
            return {"success": True, "error": None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_features(self, part_name: str) -> dict[str, Any]:
        """Return the timeline of features on the part (name + type)."""
        try:
            entry = self._get_doc_entry(part_name)
            comp_def = entry["doc"].ComponentDefinition
            feats = []
            for i in range(1, comp_def.Features.Count + 1):
                f = comp_def.Features.Item(i)
                feats.append({
                    "index": i,
                    "name": f.Name,
                    "type": f.__class__.__name__ if hasattr(f, "__class__") else "Feature",
                })
            return {"success": True, "error": None, "count": len(feats), "features": feats}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_faces(self, part_name: str) -> dict[str, Any]:
        """Return faces of body 1 with a representative point on each face (mm).

        Uses Face.PointOnFace which is more reliably exposed than RangeBox
        in pywin32 late-binding. The reported point is *on* the face, not the
        bounding-box centroid — sufficient for face-picking heuristics.
        """
        try:
            entry = self._get_doc_entry(part_name)
            comp_def = entry["doc"].ComponentDefinition
            if comp_def.SurfaceBodies.Count == 0:
                return {"success": True, "error": None, "count": 0, "faces": []}
            body = comp_def.SurfaceBodies.Item(1)
            faces = []
            for i in range(1, body.Faces.Count + 1):
                face = body.Faces.Item(i)
                # Try PointOnFace first (most reliable). Fallback to Evaluator.
                pt = None
                try:
                    pt = face.PointOnFace
                except Exception:
                    try:
                        pt = face.Evaluator.GetPointAtParam(0.5, 0.5)
                    except Exception:
                        pt = None
                if pt is not None:
                    coords_cm = (pt.X, pt.Y, pt.Z)
                    point_mm = [c * 10.0 for c in coords_cm]
                else:
                    point_mm = [None, None, None]
                faces.append({
                    "index": i,
                    "point_mm": point_mm,
                    "surface_type": getattr(face, "SurfaceType", None),
                })
            return {"success": True, "error": None, "count": len(faces), "faces": faces}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_parameters(self, part_name: str) -> dict[str, Any]:
        """Return all user parameters on the part."""
        try:
            entry = self._get_doc_entry(part_name)
            ups = entry["doc"].ComponentDefinition.Parameters.UserParameters
            params = []
            for i in range(1, ups.Count + 1):
                p = ups.Item(i)
                try:
                    val = p.Value
                except Exception:
                    val = None
                params.append({"name": p.Name, "value": val, "expression": p.Expression})
            return {"success": True, "error": None, "count": len(params), "parameters": params}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Assembly support (for multi-part designs)
    # ------------------------------------------------------------------

    def new_assembly(self, asm_name: str, template: str | None = None) -> dict[str, Any]:
        """Create a new empty AssemblyDocument registered under asm_name."""
        try:
            app = self._get_app()
            tpl = template or _DEFAULT_IAM_TEMPLATE
            if not Path(tpl).exists():
                # Fall back to FileManager.GetTemplateFile path
                try:
                    tpl = app.FileManager.GetTemplateFile(_kAssemblyDocumentObject, 26214, 65792)
                except Exception:
                    raise RuntimeError(f"assembly template not found at {tpl}")
            doc = app.Documents.Add(_kAssemblyDocumentObject, tpl, True)
            # Cast generic Document -> AssemblyDocument so ComponentDefinition
            # and other AssemblyDocument-specific properties are available
            try:
                doc = win32com.client.CastTo(doc, "AssemblyDocument")
            except Exception:
                pass  # late binding may already expose all members
            self._docs[asm_name] = {"doc": doc, "sketch": None, "is_assembly": True}
            return {"success": True, "error": None, "name": asm_name,
                    "file_type": "AssemblyDocument"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def place_component(
        self,
        asm_name: str,
        part_path: str,
        position_mm: tuple[float, float, float] = (0.0, 0.0, 0.0),
    ) -> dict[str, Any]:
        """Place a .ipt as an occurrence in the assembly at a translation offset."""
        try:
            entry = self._get_doc_entry(asm_name)
            if not entry.get("is_assembly"):
                raise RuntimeError(f"{asm_name!r} is not an assembly")
            asm_doc = entry["doc"]
            app = self._get_app()
            tg = app.TransientGeometry
            x = self._mm_to_cm(position_mm[0])
            y = self._mm_to_cm(position_mm[1])
            z = self._mm_to_cm(position_mm[2])
            matrix = tg.CreateMatrix()
            matrix.SetTranslation(tg.CreateVector(x, y, z))
            occ = asm_doc.ComponentDefinition.Occurrences.Add(
                str(Path(part_path).resolve()).replace("/", "\\"), matrix
            )
            return {"success": True, "error": None,
                    "occurrence_name": occ.Name, "part_path": part_path}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def save_assembly(self, asm_name: str, file_path: str) -> dict[str, Any]:
        """Save the assembly to a .iam file."""
        try:
            entry = self._get_doc_entry(asm_name)
            if not entry.get("is_assembly"):
                raise RuntimeError(f"{asm_name!r} is not an assembly")
            asm_doc = entry["doc"]
            out = Path(file_path).resolve()
            out.parent.mkdir(parents=True, exist_ok=True)
            asm_doc.SaveAs(str(out).replace("/", "\\"), False)
            return {"success": True, "error": None, "iam_path": str(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Assembly constraints (ground / mate / flush)
    # ------------------------------------------------------------------

    _AXIS_INDEX = {"X": 1, "Y": 2, "Z": 3}
    _PLANE_INDEX_OCC = {"YZ": 1, "XZ": 2, "XY": 3}

    def ground_component(
        self,
        asm_name: str,
        occ_name: str,
    ) -> dict[str, Any]:
        """Pin an occurrence in space so other components constrain against it.

        Convention: ground exactly one component (usually the base/housing) and
        constrain everything else relative to it.
        """
        code = (
            f'occ = comp_def.Occurrences.ItemByName("{occ_name}")\n'
            "occ.Grounded = True\n"
            'result = {"grounded": True, "occ_name": occ.Name}\n'
        )
        return self.run_python(code, part_name=asm_name)

    def assemble_axis_mate(
        self,
        asm_name: str,
        occ1_name: str,
        occ2_name: str,
        axis1: str = "Y",
        axis2: str = "Y",
    ) -> dict[str, Any]:
        """Mate two occurrences' work axes (concentric / coaxial constraint).

        Locks 4 DOF (X+Z translation, two pivot rotations); leaves rotation
        about the shared axis free — correct for a shaft.
        """
        a1 = axis1.upper()
        a2 = axis2.upper()
        if a1 not in self._AXIS_INDEX or a2 not in self._AXIS_INDEX:
            return {"success": False, "error": "axis must be X, Y or Z"}
        i1, i2 = self._AXIS_INDEX[a1], self._AXIS_INDEX[a2]
        code = (
            "import win32com.client as wc\n"
            f'occ1 = comp_def.Occurrences.ItemByName("{occ1_name}")\n'
            f'occ2 = comp_def.Occurrences.ItemByName("{occ2_name}")\n'
            "cd1 = wc.CastTo(occ1.Definition, 'PartComponentDefinition')\n"
            "cd2 = wc.CastTo(occ2.Definition, 'PartComponentDefinition')\n"
            f"ax1 = cd1.WorkAxes.Item({i1})\n"
            f"ax2 = cd2.WorkAxes.Item({i2})\n"
            "try:\n"
            "    p1 = occ1.CreateGeometryProxy(ax1)\n"
            "    p2 = occ2.CreateGeometryProxy(ax2)\n"
            "except TypeError:\n"
            "    import pythoncom\n"
            "    v1 = pythoncom.VARIANT(pythoncom.VT_DISPATCH | pythoncom.VT_BYREF, None)\n"
            "    v2 = pythoncom.VARIANT(pythoncom.VT_DISPATCH | pythoncom.VT_BYREF, None)\n"
            "    occ1.CreateGeometryProxy(ax1, v1)\n"
            "    occ2.CreateGeometryProxy(ax2, v2)\n"
            "    p1, p2 = v1.value, v2.value\n"
            "c = comp_def.Constraints.AddMateConstraint(p1, p2, 0.0)\n"
            'result = {"constraint_name": c.Name, "type": "axis_mate"}\n'
        )
        return self.run_python(code, part_name=asm_name)

    def assemble_plane_mate(
        self,
        asm_name: str,
        occ1_name: str,
        occ2_name: str,
        plane1: str = "XZ",
        plane2: str = "XZ",
        offset_mm: float = 0.0,
        flush: bool = False,
    ) -> dict[str, Any]:
        """Mate (faces opposing) or Flush (faces aligned) constraint between
        two occurrences' origin work planes, with a signed offset.

        Use after an axis mate to lock the remaining axial DOF. Origin-plane
        indices: YZ=1, XZ=2, XY=3.
        """
        p1 = plane1.upper()
        p2 = plane2.upper()
        if p1 not in self._PLANE_INDEX_OCC or p2 not in self._PLANE_INDEX_OCC:
            return {"success": False, "error": "plane must be XY, XZ or YZ"}
        i1, i2 = self._PLANE_INDEX_OCC[p1], self._PLANE_INDEX_OCC[p2]
        offset_cm = self._mm_to_cm(offset_mm)
        method = "AddFlushConstraint" if flush else "AddMateConstraint"
        kind = "flush" if flush else "mate"
        code = (
            "import win32com.client as wc\n"
            f'occ1 = comp_def.Occurrences.ItemByName("{occ1_name}")\n'
            f'occ2 = comp_def.Occurrences.ItemByName("{occ2_name}")\n'
            "cd1 = wc.CastTo(occ1.Definition, 'PartComponentDefinition')\n"
            "cd2 = wc.CastTo(occ2.Definition, 'PartComponentDefinition')\n"
            f"pl1 = cd1.WorkPlanes.Item({i1})\n"
            f"pl2 = cd2.WorkPlanes.Item({i2})\n"
            "try:\n"
            "    q1 = occ1.CreateGeometryProxy(pl1)\n"
            "    q2 = occ2.CreateGeometryProxy(pl2)\n"
            "except TypeError:\n"
            "    import pythoncom\n"
            "    v1 = pythoncom.VARIANT(pythoncom.VT_DISPATCH | pythoncom.VT_BYREF, None)\n"
            "    v2 = pythoncom.VARIANT(pythoncom.VT_DISPATCH | pythoncom.VT_BYREF, None)\n"
            "    occ1.CreateGeometryProxy(pl1, v1)\n"
            "    occ2.CreateGeometryProxy(pl2, v2)\n"
            "    q1, q2 = v1.value, v2.value\n"
            f"c = comp_def.Constraints.{method}(q1, q2, {offset_cm})\n"
            'result = {"constraint_name": c.Name, '
            f'"type": "{kind}", "offset_mm": {offset_mm}}}\n'
        )
        return self.run_python(code, part_name=asm_name)

    def loft(
        self,
        part_name: str,
        sketch_indices: list[int],
        operation: str = "new_body",
    ) -> dict[str, Any]:
        """Loft through 2+ sketches in order.

        sketch_indices are 1-based into comp_def.Sketches (the order in
        which they were created). All sketches must contain a closed
        planar profile.
        """
        if operation not in _EXTENT_OP:
            return {
                "success": False,
                "error": f"operation must be one of {list(_EXTENT_OP)}",
            }
        if len(sketch_indices) < 2:
            return {"success": False, "error": "loft requires at least 2 sketches"}
        op_val = _EXTENT_OP[operation]
        idx_list = list(sketch_indices)
        code = (
            "profiles = app.TransientObjects.CreateObjectCollection()\n"
            f"for i in {idx_list!r}:\n"
            "    sk = comp_def.Sketches.Item(i)\n"
            "    profiles.Add(sk.Profiles.AddForSolid())\n"
            "lofts = comp_def.Features.LoftFeatures\n"
            f"ldef = lofts.CreateLoftDefinition(profiles, {op_val})\n"
            "feat = lofts.Add(ldef)\n"
            f'result = {{"feature_name": feat.Name, "sections": {len(idx_list)}}}\n'
        )
        return self.run_python(code, part_name=part_name)
