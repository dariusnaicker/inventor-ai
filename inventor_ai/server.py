"""
server.py — FastMCP server exposing Autodesk Inventor automation tools.

Each tool is a thin wrapper around InventorAPI; no business logic lives here.
All return values are JSON strings (indent=2) for MCP protocol compatibility.

Run directly:
    python server.py

Or register in claude_desktop_config.json / mcp.json as:
    {
      "inventor": {
        "command": "python",
        "args": ["<absolute_path_to>/server.py"]
      }
    }
"""

from __future__ import annotations

import json

from fastmcp import FastMCP

from .api import InventorAPI

mcp = FastMCP("inventor")
_api = InventorAPI()


# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_connect() -> str:
    """
    Connect to a running Inventor instance or auto-launch it.

    Returns Inventor version and connection status.
    """
    result = _api.connect()
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Document management
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_new_part(part_name: str) -> str:
    """
    Create a new empty PartDocument registered under part_name.

    Args:
        part_name: Logical key used to reference this part in subsequent calls.
    """
    result = _api.new_part(part_name)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_list_open_documents() -> str:
    """
    List all documents currently open in Inventor (names and file paths).
    """
    result = _api.list_open_documents()
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_close_document(part_name: str, save: bool = False) -> str:
    """
    Close the document registered under part_name.

    Args:
        part_name: Logical key of the part to close.
        save: Whether to save the document before closing (default False).
    """
    result = _api.close_document(part_name, save)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_save_part(part_name: str, file_path: str) -> str:
    """
    Save the part document to an .ipt file at the given absolute path.

    Args:
        part_name: Logical key of the part.
        file_path: Absolute path for the output .ipt file.
    """
    result = _api.save_part(part_name, file_path)
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Sketching
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_new_sketch(part_name: str, plane: str = "XY") -> str:
    """
    Create a new 2-D sketch on one of the three origin work planes.

    Args:
        part_name: Logical key of the part.
        plane: "XY", "XZ", or "YZ" (default "XY").
    """
    result = _api.new_sketch(part_name, plane)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_draw_circle(
    part_name: str,
    center_x_mm: float,
    center_y_mm: float,
    diameter_mm: float,
) -> str:
    """
    Draw a circle on the active sketch.

    Args:
        part_name: Logical key of the part.
        center_x_mm: X coordinate of centre in millimetres.
        center_y_mm: Y coordinate of centre in millimetres.
        diameter_mm: Circle diameter in millimetres.
    """
    result = _api.draw_circle(part_name, center_x_mm, center_y_mm, diameter_mm)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_draw_rectangle(
    part_name: str,
    x_mm: float,
    y_mm: float,
    width_mm: float,
    height_mm: float,
) -> str:
    """
    Draw an axis-aligned rectangle on the active sketch.

    Args:
        part_name: Logical key of the part.
        x_mm: X coordinate of lower-left corner in mm.
        y_mm: Y coordinate of lower-left corner in mm.
        width_mm: Width in mm.
        height_mm: Height in mm.
    """
    result = _api.draw_rectangle(part_name, x_mm, y_mm, width_mm, height_mm)
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Features
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_extrude(
    part_name: str,
    distance_mm: float,
    direction: str = "positive",
    operation: str = "new_body",
) -> str:
    """
    Extrude the first profile in the active sketch by a fixed distance.

    Args:
        part_name: Logical key of the part.
        distance_mm: Extrusion depth in millimetres.
        direction: "positive", "negative", or "symmetric" (default "positive").
        operation: "new_body", "join", "cut", or "intersect" (default "new_body").
    """
    result = _api.extrude(part_name, distance_mm, direction, operation)
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Parameters
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_add_parameter(
    part_name: str,
    name: str,
    value: float,
    unit: str = "mm",
) -> str:
    """
    Add a user parameter to the part model.

    Args:
        part_name: Logical key of the part.
        name: Parameter name (valid Inventor identifier).
        value: Numeric value.
        unit: Unit string, e.g. "mm", "deg", "" (default "mm").
    """
    result = _api.add_parameter(part_name, name, value, unit)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_set_parameter(part_name: str, name: str, new_value: float) -> str:
    """
    Update an existing user parameter's value.

    Args:
        part_name: Logical key of the part.
        name: Parameter name.
        new_value: New numeric value (same unit as when created).
    """
    result = _api.set_parameter(part_name, name, new_value)
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_export_stl(
    part_name: str,
    output_path: str,
    resolution: str = "medium",
) -> str:
    """
    Export the part as a binary STL file for Rocky DEM import.

    Locates the STL translator add-in at runtime (DisplayName lookup).
    Sets ExportUnits=4 (mm) so the STL is in millimetres.

    Args:
        part_name: Logical key of the part.
        output_path: Absolute path for the output .stl file.
        resolution: "coarse", "medium", or "fine" (default "medium").
    """
    result = _api.export_stl(part_name, output_path, resolution)
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Mass properties
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_get_mass_properties(part_name: str) -> str:
    """
    Return volume (mm³), mass (kg), surface area (mm²), and centre of mass (mm).

    Args:
        part_name: Logical key of the part.
    """
    result = _api.get_mass_properties(part_name)
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# High-level geometry builders
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_create_cylinder(
    diameter_mm: float,
    height_mm: float,
    part_name: str,
    output_dir: str,
) -> str:
    """
    Build a solid cylinder, save as .ipt, and export as .stl in one call.

    Used to generate oedometer container geometry for Rocky DEM.

    Args:
        diameter_mm: Outer diameter in mm.
        height_mm: Total height in mm.
        part_name: Stem name for output files (e.g. "oedometer_cylinder").
        output_dir: Absolute path to the directory for .ipt and .stl files.
    """
    result = _api.create_cylinder(diameter_mm, height_mm, part_name, output_dir)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_create_box(
    width_mm: float,
    depth_mm: float,
    height_mm: float,
    part_name: str,
    output_dir: str,
) -> str:
    """
    Build a solid rectangular box, save as .ipt, and export as .stl in one call.

    Used to generate shear cell / direct shear box geometry for Rocky DEM.

    Args:
        width_mm: X dimension in mm.
        depth_mm: Y dimension in mm.
        height_mm: Z dimension in mm.
        part_name: Stem name for output files (e.g. "shear_box").
        output_dir: Absolute path to the directory for .ipt and .stl files.
    """
    result = _api.create_box(width_mm, depth_mm, height_mm, part_name, output_dir)
    return json.dumps(result, indent=2)


@mcp.tool()
def inventor_create_funnel(
    top_diameter_mm: float,
    bottom_diameter_mm: float,
    height_mm: float,
    part_name: str,
    output_dir: str,
) -> str:
    """
    Build a truncated-cone funnel, save as .ipt, and export as .stl in one call.

    Used to generate angle-of-repose funnel geometry for Rocky DEM.
    Geometry is a revolution of a trapezoidal half-profile around the Y-axis.

    Args:
        top_diameter_mm: Diameter of the wide top opening in mm.
        bottom_diameter_mm: Diameter of the narrow bottom outlet in mm.
        height_mm: Axial height of the funnel in mm.
        part_name: Stem name for output files (e.g. "aor_funnel").
        output_dir: Absolute path to the directory for .ipt and .stl files.
    """
    result = _api.create_funnel(
        top_diameter_mm, bottom_diameter_mm, height_mm, part_name, output_dir
    )
    return json.dumps(result, indent=2)


# ---------------------------------------------------------------------------
# Diagnostics, escape hatch, and advanced features (loft/shell/fillet)
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_test_connection() -> str:
    """Ping Inventor and report version, visibility, and open document count.

    Use as a smoke test before any other tool call.
    """
    return json.dumps(_api.test_connection(), indent=2)


@mcp.tool()
def inventor_run_python(code: str, part_name: str | None = None) -> str:
    """Execute arbitrary Python against Inventor's COM API (escape hatch).

    Pre-injected names available in the code's namespace:
        app, doc, comp_def, sketch, tg, mm_to_cm,
        EXTENT_OP, EXTENT_DIR, PLANE_INDEX, api

    Set a variable named ``result`` in the code to return a value.
    Use this when no dedicated tool exists for the desired operation
    (e.g. mirror, pattern, sweep, draft, complex face selection).

    Args:
        code: Python source.
        part_name: optional logical part name to wire up doc/comp_def/sketch.
    """
    return json.dumps(_api.run_python(code, part_name=part_name), indent=2)


@mcp.tool()
def inventor_fillet_all_edges(part_name: str, radius_mm: float) -> str:
    """Round every edge of the part's first body with a uniform radius."""
    return json.dumps(_api.fillet_all_edges(part_name, radius_mm), indent=2)


@mcp.tool()
def inventor_shell(part_name: str, thickness_mm: float, face_filter: str = "top") -> str:
    """Hollow the part's first body, removing one face to leave a cavity.

    Args:
        face_filter: which face to remove. One of "top", "bottom", "+z", "-z".
    """
    return json.dumps(_api.shell(part_name, thickness_mm, face_filter), indent=2)


@mcp.tool()
def inventor_loft(
    part_name: str,
    sketch_indices: list[int],
    operation: str = "new_body",
) -> str:
    """Loft through 2+ sketches in order.

    Args:
        sketch_indices: 1-based indices into comp_def.Sketches.
        operation: "new_body", "join", "cut", or "intersect".
    """
    return json.dumps(_api.loft(part_name, sketch_indices, operation), indent=2)


@mcp.tool()
def inventor_sweep(
    part_name: str,
    profile_sketch_idx: int,
    path_sketch_idx: int,
    operation: str = "new_body",
) -> str:
    """Sweep a profile sketch along a path sketch (both 1-based)."""
    return json.dumps(
        _api.sweep(part_name, profile_sketch_idx, path_sketch_idx, operation), indent=2
    )


@mcp.tool()
def inventor_mirror(part_name: str, plane: str = "XY") -> str:
    """Mirror the most recent feature across an origin plane (XY/XZ/YZ)."""
    return json.dumps(_api.mirror(part_name, plane), indent=2)


@mcp.tool()
def inventor_circular_pattern(
    part_name: str, count: int, axis: str = "Y", angle_deg: float = 360.0
) -> str:
    """Circular-pattern the most recent feature around an origin axis (X/Y/Z)."""
    return json.dumps(
        _api.circular_pattern(part_name, count, axis, angle_deg), indent=2
    )


@mcp.tool()
def inventor_rectangular_pattern(
    part_name: str,
    count_x: int,
    count_y: int,
    spacing_x_mm: float,
    spacing_y_mm: float,
    axis_x: str = "X",
    axis_y: str = "Z",
) -> str:
    """Rectangular-pattern the most recent feature in two directions."""
    return json.dumps(
        _api.rectangular_pattern(
            part_name, count_x, count_y, spacing_x_mm, spacing_y_mm, axis_x, axis_y
        ),
        indent=2,
    )


@mcp.tool()
def inventor_export_step(part_name: str, output_path: str) -> str:
    """Export the part as a STEP AP214 file."""
    return json.dumps(_api.export_step(part_name, output_path), indent=2)


@mcp.tool()
def inventor_undo() -> str:
    """Send a single Undo command to Inventor."""
    return json.dumps(_api.undo(), indent=2)


@mcp.tool()
def inventor_list_features(part_name: str) -> str:
    """List the timeline of features on the part (name + index)."""
    return json.dumps(_api.list_features(part_name), indent=2)


@mcp.tool()
def inventor_list_faces(part_name: str) -> str:
    """List faces of body 1 with their bounding-box centroid in mm."""
    return json.dumps(_api.list_faces(part_name), indent=2)


@mcp.tool()
def inventor_list_parameters(part_name: str) -> str:
    """List all user parameters on the part."""
    return json.dumps(_api.list_parameters(part_name), indent=2)


# ---------------------------------------------------------------------------
# Assembly support
# ---------------------------------------------------------------------------

@mcp.tool()
def inventor_new_assembly(asm_name: str, template: str | None = None) -> str:
    """Create a new empty AssemblyDocument under asm_name."""
    return json.dumps(_api.new_assembly(asm_name, template), indent=2)


@mcp.tool()
def inventor_place_component(
    asm_name: str,
    part_path: str,
    position_mm: tuple[float, float, float] = (0.0, 0.0, 0.0),
) -> str:
    """Place a .ipt as an occurrence in the assembly at a translation offset (mm)."""
    return json.dumps(
        _api.place_component(asm_name, part_path, position_mm), indent=2
    )


@mcp.tool()
def inventor_save_assembly(asm_name: str, file_path: str) -> str:
    """Save the assembly to a .iam file."""
    return json.dumps(_api.save_assembly(asm_name, file_path), indent=2)


@mcp.tool()
def inventor_ground_component(asm_name: str, occ_name: str) -> str:
    """Pin an occurrence in space as the assembly's reference frame."""
    return json.dumps(_api.ground_component(asm_name, occ_name), indent=2)


@mcp.tool()
def inventor_assemble_axis_mate(
    asm_name: str,
    occ1_name: str,
    occ2_name: str,
    axis1: str = "Y",
    axis2: str = "Y",
) -> str:
    """Mate two occurrences' work axes (concentric / coaxial). Locks 4 DOF."""
    return json.dumps(
        _api.assemble_axis_mate(asm_name, occ1_name, occ2_name, axis1, axis2),
        indent=2,
    )


@mcp.tool()
def inventor_assemble_plane_mate(
    asm_name: str,
    occ1_name: str,
    occ2_name: str,
    plane1: str = "XZ",
    plane2: str = "XZ",
    offset_mm: float = 0.0,
    flush: bool = False,
) -> str:
    """Mate (faces opposing) or Flush (faces aligned) two work planes with offset."""
    return json.dumps(
        _api.assemble_plane_mate(
            asm_name, occ1_name, occ2_name, plane1, plane2, offset_mm, flush
        ),
        indent=2,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
