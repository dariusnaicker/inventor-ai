"""Build a simplified turbocharger cartridge as a fully constrained Inventor assembly.

Demonstrates the full Inventor MCP capability stack including assembly constraints
on a real-world part with parts that actually mate physically.

Parts produced (saved to rocky/geometry/turbo_demo/):
  1. turbo_shaft.ipt              stepped shaft, turbine at -Y, compressor at +Y
  2. turbo_bearing_housing.ipt    revolved profile: 30mm bore lower, 70mm cavity upper
  3. turbo_compressor_wheel.ipt   hub with 26mm bore + 8 blades
  4. turbocharger.iam             fully constrained assembly

Assembly layout (along global Y axis, origin at housing bottom face):
    Y =  -60 .. -20 .. 0 .. 60 ......... 100 .. 120
         [turb][middle][bore | cavity      ][shaft tip]
              shaft passes through bore (30mm) and into cavity
                              wheel sits on shaft compressor end inside cavity

Usage (Inventor 2026 must be running):
    uv run python scripts/build_turbo_demo.py
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from inventor_ai.api import InventorAPI, _EXTENT_OP  # noqa: E402

OUT = REPO_ROOT / "_turbo_demo_out"
OUT.mkdir(parents=True, exist_ok=True)


def _check(label: str, result: dict) -> None:
    """Halt the build if a step failed; otherwise log success."""
    if not result.get("success"):
        print(f"  [FAIL] {label}: {result.get('error')}")
        raise SystemExit(1)
    print(f"  [OK ] {label}")


def revolve_profile(api: InventorAPI, part_name: str, points_mm: list[tuple[float, float]]) -> None:
    """Revolve a closed polyline (radius, height) around the Y axis."""
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


# ---------------------------------------------------------------------------
# Cleanup: close any stale documents from a previous run
# ---------------------------------------------------------------------------
def cleanup_open_docs(api: InventorAPI) -> None:
    print("\n=== closing any stale Inventor documents ===")
    code = """
to_close = []
for i in range(1, app.Documents.Count + 1):
    d = app.Documents.Item(i)
    # 12290 = kPartDocumentObject, 12291 = kAssemblyDocumentObject
    if d.DocumentType in (12290, 12291):
        to_close.append(d)
closed = 0
for d in list(to_close):
    try:
        d.Close(True)  # SkipSave = True
        closed += 1
    except Exception:
        pass
result = {"closed": closed}
"""
    res = api.run_python(code, part_name=None)
    if res.get("success"):
        n = res.get("result", {}).get("closed", 0)
        print(f"  [OK ] closed {n} stale documents")
    # Always wipe the internal registry — stale entries would block re-creation
    api._docs.clear()


# ---------------------------------------------------------------------------
# Part 1: turbo_shaft.ipt — a stepped shaft
# ---------------------------------------------------------------------------
def build_shaft(api: InventorAPI) -> Path:
    """Stepped shaft (turbine at -Y, compressor at +Y).

    Y_local 0-40    turbine end    20 mm dia
    Y_local 40-120  middle         30 mm dia (rides in bearing housing bore)
    Y_local 120-180 compressor end 24 mm dia (carries the wheel)
    """
    name = "turbo_shaft"
    print(f"\n=== {name}: stepped shaft (revolve) ===")
    profile = [
        (0.0,  0.0),
        (10.0, 0.0),
        (10.0, 40.0),    # turbine end -> middle step
        (15.0, 40.0),
        (15.0, 120.0),   # middle -> compressor step
        (12.0, 120.0),
        (12.0, 180.0),   # 24 mm dia compressor end
        (0.0,  180.0),
    ]
    revolve_profile(api, name, profile)
    print(f"  [OK ] revolve 8-point stepped profile")

    _check("fillet 1 mm on all edges", api.fillet_all_edges(name, 1.0))
    out = OUT / f"{name}.ipt"
    _check("save_part", api.save_part(name, str(out)))
    _check("export_stl", api.export_stl(name, str(OUT / f"{name}.stl")))
    return out


# ---------------------------------------------------------------------------
# Part 2: turbo_bearing_housing.ipt — cylinder with central bore + 6 bolt holes
# ---------------------------------------------------------------------------
def build_bearing_housing(api: InventorAPI) -> Path:
    """Stepped housing built as a single revolved profile: 30 mm bore from
    Y=0 to Y=60 (carries the shaft), 70 mm cavity from Y=60 to Y=100 (holds
    the compressor wheel). 90 mm OD overall, 100 mm tall. 6 bolt holes
    through the bottom flange on a 60 mm pitch circle.
    """
    name = "turbo_bearing_housing"
    print(f"\n=== {name}: revolved bore+cavity body + bolt circle ===")

    # Cross-section (R, Y) revolved 360 deg around Y axis.
    profile = [
        (15.0,   0.0),   # bore wall, bottom face
        (15.0,  60.0),   # bore wall, top
        (35.0,  60.0),   # ledge at bore/cavity transition
        (35.0, 100.0),   # cavity wall, top
        (45.0, 100.0),   # outer wall, top
        (45.0,   0.0),   # outer wall, bottom
    ]
    revolve_profile(api, name, profile)
    print(f"  [OK ] revolve 6-point housing profile (bore + cavity)")

    code = """
import math
pitch_radius_cm = 3.0   # 60/2 mm
hole_radius_cm = 0.3    # 6/2 mm
sketch = comp_def.Sketches.Add(comp_def.WorkPlanes.Item(2))  # XZ at Y=0
for i in range(6):
    theta = i * (2 * math.pi / 6)
    cx = pitch_radius_cm * math.cos(theta)
    cz = pitch_radius_cm * math.sin(theta)
    sketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(cx, cz), hole_radius_cm)
profile = sketch.Profiles.AddForSolid()
extrudes = comp_def.Features.ExtrudeFeatures
ext_def = extrudes.CreateExtrudeDefinition(profile, EXTENT_OP['cut'])
# Cut 10 mm deep through the bottom flange (only where there's material at R=30)
ext_def.SetDistanceExtent(1.0, EXTENT_DIR['positive'])
feat = extrudes.Add(ext_def)
result = {"feature_name": feat.Name, "holes": 6}
"""
    _check("cut 6 bolt holes (run_python loop)", api.run_python(code, part_name=name))
    _check("fillet 1 mm on all edges", api.fillet_all_edges(name, 1.0))

    out = OUT / f"{name}.ipt"
    _check("save_part", api.save_part(name, str(out)))
    _check("export_stl", api.export_stl(name, str(OUT / f"{name}.stl")))
    _check("export_step", api.export_step(name, str(OUT / f"{name}.step")))
    return out


# ---------------------------------------------------------------------------
# Part 3: turbo_compressor_wheel.ipt — hub + 8 blades
# ---------------------------------------------------------------------------
def build_compressor_wheel(api: InventorAPI) -> Path:
    """Compressor wheel: revolved hub with a 26 mm dia center bore so it can
    sit on the shaft compressor end. 8 radial blades extruded around it.

    Exducer (wide, 80 mm dia) at Y=0, inducer (15 mm radius) at Y=30.
    Center bore (26 mm dia, R=13) runs the full hub height for shaft mount.
    """
    name = "turbo_compressor_wheel"
    print(f"\n=== {name}: hub with center bore + 8 blades ===")

    # Hub cross-section: closed loop in (R, Y) space, NOT touching axis (bore).
    hub_profile = [
        (13.0,  0.0),    # bore wall, exducer face
        (13.0, 30.0),    # bore wall, inducer face
        (15.0, 30.0),    # inducer outer top
        (25.0, 20.0),    # mid-shroud cone
        (40.0,  5.0),    # exducer outer
        (40.0,  0.0),    # exducer face outer corner
    ]
    revolve_profile(api, name, hub_profile)
    print(f"  [OK ] revolve hub with center bore")

    code = """
import math
n_blades = 8
blade_h_mm = 15.0
blade_thick_mm = 2.0
base_r_mm = 14.0   # outside the 13 mm bore wall
tip_r_mm = 40.0
base_y_mm = 15.0
for i in range(n_blades):
    theta = i * (2 * math.pi / n_blades)
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)
    corners_mm = [
        (base_r_mm, -blade_thick_mm/2),
        (tip_r_mm,  -blade_thick_mm/2),
        (tip_r_mm,   blade_thick_mm/2),
        (base_r_mm,  blade_thick_mm/2),
    ]
    rotated_cm = []
    for x_mm, z_mm in corners_mm:
        rx = x_mm * cos_t - z_mm * sin_t
        rz = x_mm * sin_t + z_mm * cos_t
        rotated_cm.append((mm_to_cm(rx), mm_to_cm(rz)))
    sketch = comp_def.Sketches.Add(comp_def.WorkPlanes.Item(2))  # XZ
    lines = sketch.SketchLines
    first = lines.AddByTwoPoints(
        tg.CreatePoint2d(*rotated_cm[0]),
        tg.CreatePoint2d(*rotated_cm[1]),
    )
    prev = first
    for pt in rotated_cm[2:]:
        prev = lines.AddByTwoPoints(prev.EndSketchPoint, tg.CreatePoint2d(*pt))
    lines.AddByTwoPoints(prev.EndSketchPoint, first.StartSketchPoint)
    profile = sketch.Profiles.AddForSolid()
    extrudes = comp_def.Features.ExtrudeFeatures
    ext_def = extrudes.CreateExtrudeDefinition(profile, EXTENT_OP['join'])
    ext_def.SetDistanceExtent(mm_to_cm(blade_h_mm + base_y_mm), EXTENT_DIR['positive'])
    extrudes.Add(ext_def)
result = {"blades_added": n_blades}
"""
    _check("extrude 8 blades (run_python loop)", api.run_python(code, part_name=name))

    out = OUT / f"{name}.ipt"
    _check("save_part", api.save_part(name, str(out)))
    _check("export_stl", api.export_stl(name, str(OUT / f"{name}.stl")))
    return out


# ---------------------------------------------------------------------------
# Assembly: turbocharger.iam — fully constrained
# ---------------------------------------------------------------------------
def build_assembly(api: InventorAPI, shaft: Path, housing: Path, wheel: Path) -> Path:
    """Build the .iam, place 3 occurrences, ground the housing, axis-mate
    shaft + wheel to housing (concentric on Y), and lock axial position with
    flush plane mates. Final assembly has 1 free DOF (shaft rotation).

    Geometry summary (assembly coords, Y axis):
      housing: Y=0..100   bore Y=0..60 (R=15), cavity Y=60..100 (R=35)
      shaft:   Y=-60..120 turbine -60..-20, middle -20..60 (in bore),
                          compressor 60..120 (in cavity, sticks 20mm out top)
      wheel:   Y=65..95   hub centered in cavity, 26 mm bore on shaft tip
    """
    print("\n=== turbocharger.iam: constrained assembly ===")
    asm_name = "turbocharger"
    _check("new_assembly", api.new_assembly(asm_name))

    h_res = api.place_component(asm_name, str(housing), (0.0, 0.0, 0.0))
    _check("place housing", h_res)
    h_occ = h_res["occurrence_name"]
    _check("ground housing", api.ground_component(asm_name, h_occ))

    # Shaft: -60 mm Y so middle (local 40..120) spans bore (assembly 0..60),
    # compressor end (local 120..180) sits in cavity (assembly 60..120).
    s_res = api.place_component(asm_name, str(shaft), (0.0, -60.0, 0.0))
    _check("place shaft", s_res)
    s_occ = s_res["occurrence_name"]
    _check(
        "shaft axis-mate to housing (Y)",
        api.assemble_axis_mate(asm_name, h_occ, s_occ),
    )
    _check(
        "shaft XZ plane flush -60 mm",
        api.assemble_plane_mate(
            asm_name, h_occ, s_occ, offset_mm=-60.0, flush=True,
        ),
    )

    # Wheel: +65 mm Y so hub (local 0..30) sits inside the cavity (Y=65..95),
    # bore around the shaft compressor end. 5 mm clearance to cavity floor and
    # cavity ceiling.
    w_res = api.place_component(asm_name, str(wheel), (0.0, 65.0, 0.0))
    _check("place compressor wheel", w_res)
    w_occ = w_res["occurrence_name"]
    _check(
        "wheel axis-mate to housing (Y)",
        api.assemble_axis_mate(asm_name, h_occ, w_occ),
    )
    _check(
        "wheel XZ plane flush +65 mm",
        api.assemble_plane_mate(
            asm_name, h_occ, w_occ, offset_mm=65.0, flush=True,
        ),
    )

    out = OUT / f"{asm_name}.iam"
    _check("save_assembly", api.save_assembly(asm_name, str(out)))
    return out


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
    housing = build_bearing_housing(api)
    wheel = build_compressor_wheel(api)
    assembly = build_assembly(api, shaft, housing, wheel)

    print("\n" + "=" * 60)
    print("TURBOCHARGER DEMO BUILD COMPLETE")
    print("=" * 60)
    for f in sorted(OUT.glob("*")):
        size = f.stat().st_size
        print(f"  {f.name:35s}  {size:>12,d} B")
    print(f"\nAll outputs in: {OUT}")
    print(f"Open the assembly: {assembly}")


if __name__ == "__main__":
    main()
