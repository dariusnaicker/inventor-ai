"""IMD 254 centrifugal pump shaft, Stellenbosch University 2nd-year ME drawing project.

Real-world capability demo: build the shaft from a 2nd-year mechanical engineering
drawing brief end-to-end through Claude + the inventor-ai MCP. The brief is
private course material; only the geometry produced from the public dimensions
is committed here.

Geometry (from the IMD 254 2023 brief, all mm):
  Total length      133 (+0.0 / -0.1)
  Left keyway end   Ø12 with 15 mm keyway slot (end milled, H8)
  Seal area         Ø8  (flats on two sides, simplified to round here)
  Roller bearing B  Ø15 (cylindrical roller bearing seat, SKF NJ)
  Middle datum      Ø18 (+0.00 / -0.02)
  Ball bearing C    Ø15 (self-aligning ball bearing seat, SKF NJ self-aligning)
  Right end         Ø10 with f7 fit (drive end), 7 mm across-flats square drive
  End face          M4 blind threaded hole, 10 mm deep

Simplifications vs the full drawing:
  * Flats on Ø8 and on Ø10 right end omitted (round only)
  * Right-end keyway omitted (left keyway shown)
  * Tolerances and surface finish symbols are not modelled (those go on the
    .idw drawing, not the .ipt part)
  * Chamfers (0.5 x 45 deg) omitted

Usage (Inventor 2026 must be running):
    python examples/imd254_shaft.py
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from inventor_ai.api import InventorAPI, _EXTENT_OP  # noqa: E402

OUT = REPO_ROOT / "_imd254_shaft_out"
OUT.mkdir(parents=True, exist_ok=True)


def _check(label: str, result: dict) -> None:
    if not result.get("success"):
        print(f"  [FAIL] {label}: {result.get('error')}")
        raise SystemExit(1)
    print(f"  [OK ] {label}")


def revolve_profile(api: InventorAPI, part_name: str, points_mm: list[tuple[float, float]]) -> None:
    """Revolve a closed (R, Y) polyline around the Y axis."""
    api.new_part(part_name)
    api.new_sketch(part_name, "XY")
    entry = api._get_doc_entry(part_name)
    app = api._get_app()
    tg = app.TransientGeometry
    sketch = entry["sketch"]
    pts = [(api._mm_to_cm(r), api._mm_to_cm(h)) for r, h in points_mm]
    lines = sketch.SketchLines
    first = lines.AddByTwoPoints(tg.CreatePoint2d(*pts[0]), tg.CreatePoint2d(*pts[1]))
    prev = first
    for r, h in pts[2:]:
        prev = lines.AddByTwoPoints(prev.EndSketchPoint, tg.CreatePoint2d(r, h))
    lines.AddByTwoPoints(prev.EndSketchPoint, first.StartSketchPoint)
    profile = sketch.Profiles.AddForSolid()
    comp_def = entry["doc"].ComponentDefinition
    y_axis = comp_def.WorkAxes.Item(2)
    comp_def.Features.RevolveFeatures.AddFull(profile, y_axis, _EXTENT_OP["new_body"])


def cleanup_open_docs(api: InventorAPI) -> None:
    print("\n=== closing any stale Inventor documents ===")
    code = """
to_close = []
for i in range(1, app.Documents.Count + 1):
    d = app.Documents.Item(i)
    if d.DocumentType in (12290, 12291):
        to_close.append(d)
closed = 0
for d in list(to_close):
    try:
        d.Close(True)
        closed += 1
    except Exception:
        pass
result = {"closed": closed}
"""
    res = api.run_python(code, part_name=None)
    if res.get("success"):
        n = res.get("result", {}).get("closed", 0)
        print(f"  [OK ] closed {n} stale documents")
    api._docs.clear()


def build_shaft(api: InventorAPI) -> Path:
    name = "imd254_shaft"
    print(f"\n=== {name}: stepped pump shaft (revolve) ===")

    # Stepped profile, axis = Y, all units mm, R = radial distance from axis.
    # Lengths stack from left (Y=0) to right (Y=133):
    #   23 + 24 + 11 + 39 + 14 + 22 = 133
    profile = [
        (0.0,    0.0),
        (6.0,    0.0),    # Ø12 keyway end (left)
        (6.0,   23.0),
        (4.0,   23.0),    # step down to Ø8 seal area
        (4.0,   47.0),
        (7.5,   47.0),    # step up to Ø15 roller bearing seat (B)
        (7.5,   58.0),
        (9.0,   58.0),    # step up to Ø18 middle datum
        (9.0,   97.0),
        (7.5,   97.0),    # step down to Ø15 ball bearing seat (C)
        (7.5,  111.0),
        (5.0,  111.0),    # step down to Ø10 f7 right end
        (5.0,  133.0),
        (0.0,  133.0),
    ]
    revolve_profile(api, name, profile)
    print(f"  [OK ] revolve 14-point stepped profile (133 mm long)")

    # Left keyway slot: 15 mm long, 3 mm wide, 3 mm deep, end-milled, on the
    # +X face of the Ø12 section. Centred at Y=15.
    # Sketch on a workplane offset to the shaft surface (X=6), then cut-extrude
    # toward the axis (-X) by 3 mm so the slot has the correct flat-bottom
    # depth instead of going through the whole half-shaft.
    code = """
slot_len_mm = 15.0
slot_w_mm = 3.0
slot_depth_mm = 3.0
slot_y_centre_mm = 15.0
shaft_surface_x_cm = mm_to_cm(6.0)

# Workplane parallel to YZ, offset to the +X surface of the shaft.
yz_origin = comp_def.WorkPlanes.Item(1)
surface_plane = comp_def.WorkPlanes.AddByPlaneAndOffset(yz_origin, shaft_surface_x_cm)
surface_plane.Visible = False

sketch = comp_def.Sketches.Add(surface_plane)
half_w_cm = mm_to_cm(slot_w_mm / 2)
y_lo_cm = mm_to_cm(slot_y_centre_mm - slot_len_mm/2 + slot_w_mm/2)
y_hi_cm = mm_to_cm(slot_y_centre_mm + slot_len_mm/2 - slot_w_mm/2)

sk_arcs = sketch.SketchArcs
sk_lines = sketch.SketchLines
arc_lo = sk_arcs.AddByCenterStartEndPoint(
    tg.CreatePoint2d(y_lo_cm, 0.0),
    tg.CreatePoint2d(y_lo_cm, -half_w_cm),
    tg.CreatePoint2d(y_lo_cm,  half_w_cm),
    True,
)
arc_hi = sk_arcs.AddByCenterStartEndPoint(
    tg.CreatePoint2d(y_hi_cm, 0.0),
    tg.CreatePoint2d(y_hi_cm,  half_w_cm),
    tg.CreatePoint2d(y_hi_cm, -half_w_cm),
    True,
)
sk_lines.AddByTwoPoints(arc_lo.StartSketchPoint, arc_hi.EndSketchPoint)
sk_lines.AddByTwoPoints(arc_hi.StartSketchPoint, arc_lo.EndSketchPoint)

profile = sketch.Profiles.AddForSolid()
extrudes = comp_def.Features.ExtrudeFeatures
ext_def = extrudes.CreateExtrudeDefinition(profile, EXTENT_OP['cut'])
# Cut from surface inward (toward axis, -X) by 3 mm.
ext_def.SetDistanceExtent(mm_to_cm(slot_depth_mm), EXTENT_DIR['negative'])
feat = extrudes.Add(ext_def)
result = {"feature_name": feat.Name, "slot_length_mm": slot_len_mm}
"""
    _check("cut keyway slot 15x3x3 mm on Ø12 end", api.run_python(code, part_name=name))

    # M4 blind threaded hole at right end face (Y=133), 10 mm deep, drilled
    # along -Y from the right face. Modelled as a Ø3.3 (M4 tapping drill) hole.
    code = """
import math
hole_dia_mm = 3.3   # M4 tapping drill
hole_depth_mm = 10.0
right_face_y_cm = mm_to_cm(133.0)

# Sketch on XZ plane offset to the right face. Use a work plane: XZ origin
# plane (index 2) is at Y=0; we need a plane at Y=133. Create offset plane.
xz_origin = comp_def.WorkPlanes.Item(2)
work_planes = comp_def.WorkPlanes
right_face_plane = work_planes.AddByPlaneAndOffset(xz_origin, right_face_y_cm)
right_face_plane.Visible = False

sketch = comp_def.Sketches.Add(right_face_plane)
sketch.SketchCircles.AddByCenterRadius(
    tg.CreatePoint2d(0.0, 0.0), mm_to_cm(hole_dia_mm/2)
)
profile = sketch.Profiles.AddForSolid()
extrudes = comp_def.Features.ExtrudeFeatures
ext_def = extrudes.CreateExtrudeDefinition(profile, EXTENT_OP['cut'])
# Cut in -Y direction from the right face by 10 mm
ext_def.SetDistanceExtent(mm_to_cm(hole_depth_mm), EXTENT_DIR['negative'])
feat = extrudes.Add(ext_def)
result = {"feature_name": feat.Name, "hole_depth_mm": hole_depth_mm}
"""
    _check("drill M4 blind hole 10 mm deep at right end face",
           api.run_python(code, part_name=name))

    out_ipt = OUT / f"{name}.ipt"
    _check("save_part",   api.save_part(name, str(out_ipt)))
    _check("export_stl",  api.export_stl(name, str(OUT / f"{name}.stl")))
    _check("export_step", api.export_step(name, str(OUT / f"{name}.step")))
    return out_ipt


def main() -> None:
    api = InventorAPI()
    print("Connecting to Inventor 2026...")
    conn = api.connect()
    if not conn.get("success"):
        print(f"FATAL: {conn.get('error')}")
        sys.exit(1)
    print(f"Connected: {conn.get('version')}")

    cleanup_open_docs(api)
    shaft = build_shaft(api)

    print("\n" + "=" * 60)
    print("IMD 254 SHAFT BUILD COMPLETE")
    print("=" * 60)
    for f in sorted(OUT.glob("*")):
        if f.is_file():
            print(f"  {f.name:30s}  {f.stat().st_size:>12,d} B")
    print(f"\nAll outputs in: {OUT}")
    print(f"Open: {shaft}")


if __name__ == "__main__":
    main()
