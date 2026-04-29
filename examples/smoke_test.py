"""Smoke test for the InventorAPI methods.

Exercises every method on a simple part and a few feature-test parts. Reports
OK/FAIL per method with the actual error string so COM API issues are visible.

Usage (Inventor 2026 must be running):
    python examples/smoke_test.py
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from inventor_ai.api import InventorAPI  # noqa: E402

OUT = REPO_ROOT / "_smoke_test_out"
OUT.mkdir(parents=True, exist_ok=True)


def test(label: str, fn) -> bool:
    try:
        result = fn()
        ok = result.get("success", False)
        marker = "OK  " if ok else "FAIL"
        print(f"  [{marker}] {label}: {result}")
        return ok
    except Exception as e:
        print(f"  [FAIL] {label}: EXCEPTION {type(e).__name__}: {e}")
        return False


def main() -> None:
    api = InventorAPI()
    print("=== InventorAPI smoke test ===\n")

    # 1. test_connection
    print("Section 1: connection")
    if not test("test_connection", lambda: api.test_connection()):
        print("FATAL: cannot connect to Inventor — is it running?")
        sys.exit(1)

    # 2. Simple part: a 50 x 30 x 20 mm box
    print("\nSection 2: simple box part")
    test("new_part(testbox)",     lambda: api.new_part("testbox"))
    test("new_sketch(XY)",        lambda: api.new_sketch("testbox", "XY"))
    test("draw_rectangle 50x30",  lambda: api.draw_rectangle("testbox", -25.0, -15.0, 25.0, 15.0))
    test("extrude 20mm",          lambda: api.extrude("testbox", 20.0))

    # 3. New methods on this box
    print("\nSection 3: new feature methods")
    test("list_features",         lambda: api.list_features("testbox"))
    test("list_faces",            lambda: api.list_faces("testbox"))
    test("fillet_all_edges 2mm",  lambda: api.fillet_all_edges("testbox", 2.0))
    test("undo (remove fillet)",  lambda: api.undo())
    # Re-fillet after undo
    test("fillet_all_edges 2mm",  lambda: api.fillet_all_edges("testbox", 2.0))

    # 4. run_python escape hatch
    print("\nSection 4: run_python escape hatch")
    code = """
result = {
    "doc_name": doc.DisplayName,
    "feature_count": comp_def.Features.Count,
    "body_count": comp_def.SurfaceBodies.Count,
    "first_body_face_count": comp_def.SurfaceBodies.Item(1).Faces.Count,
}
"""
    test("run_python(introspect)", lambda: api.run_python(code, "testbox"))

    # 5. Save part + STL + STEP
    print("\nSection 5: save and export")
    box_ipt = str(OUT / "testbox.ipt")
    test("save_part",             lambda: api.save_part("testbox", box_ipt))
    test("export_stl",            lambda: api.export_stl("testbox", str(OUT / "testbox.stl")))
    test("export_step",           lambda: api.export_step("testbox", str(OUT / "testbox.step")))

    # 6. Pattern test on a fresh part: cylinder with circular pattern of small holes
    print("\nSection 6: circular_pattern on a flange")
    test("new_part(flange)",      lambda: api.new_part("flange"))
    test("new_sketch(XY)",        lambda: api.new_sketch("flange", "XY"))
    test("draw_circle 80mm",      lambda: api.draw_circle("flange", 0.0, 0.0, 80.0))
    test("extrude 5mm",           lambda: api.extrude("flange", 5.0))
    # bolt hole near edge
    test("new_sketch(XY)",        lambda: api.new_sketch("flange", "XY"))
    test("draw_circle 6mm @ R30", lambda: api.draw_circle("flange", 30.0, 0.0, 6.0))
    test("extrude cut -5mm",      lambda: api.extrude("flange", 5.0,
                                                      direction="negative",
                                                      operation="cut"))
    test("circular_pattern 6 around Y", lambda: api.circular_pattern("flange", 6, axis="Y"))
    test("save_part flange",      lambda: api.save_part("flange",
                                                        str(OUT / "flange.ipt")))

    # 7. Mirror test: build a stub on +X side, mirror to -X
    print("\nSection 7: mirror across YZ plane")
    test("new_part(mirror_test)", lambda: api.new_part("mirror_test"))
    test("new_sketch(XY)",        lambda: api.new_sketch("mirror_test", "XY"))
    test("draw_rect base",        lambda: api.draw_rectangle("mirror_test", -50.0, -10.0, -10.0, 10.0))
    test("extrude 10mm",          lambda: api.extrude("mirror_test", 10.0))
    test("mirror across YZ",      lambda: api.mirror("mirror_test", plane="YZ"))
    test("save mirror_test",      lambda: api.save_part("mirror_test",
                                                        str(OUT / "mirror_test.ipt")))

    # 8. rectangular_pattern test: small post, pattern in X and Z
    print("\nSection 8: rectangular_pattern")
    test("new_part(rect_pat)",    lambda: api.new_part("rect_pat"))
    test("new_sketch(XZ)",        lambda: api.new_sketch("rect_pat", "XZ"))
    test("draw 100x100 base",     lambda: api.draw_rectangle("rect_pat", 0.0, 0.0, 100.0, 100.0))
    test("extrude base 5mm",      lambda: api.extrude("rect_pat", 5.0))
    test("new_sketch(XZ) post",   lambda: api.new_sketch("rect_pat", "XZ"))
    test("draw_circle 6mm @20,20", lambda: api.draw_circle("rect_pat", 20.0, 20.0, 6.0))
    test("extrude post +10mm",    lambda: api.extrude("rect_pat", 10.0, operation="join"))
    test("rect_pattern 3x3",      lambda: api.rectangular_pattern("rect_pat", 3, 3, 25.0, 25.0,
                                                                  axis_x="X", axis_y="Z"))
    test("save rect_pat",         lambda: api.save_part("rect_pat",
                                                        str(OUT / "rect_pat.ipt")))

    # 9. list_parameters with add_parameter
    print("\nSection 9: parameters round-trip")
    test("add_parameter D=80",    lambda: api.add_parameter("flange", "flange_dia", 80.0))
    test("list_parameters",       lambda: api.list_parameters("flange"))

    print("\n=== smoke test complete ===")
    print(f"Outputs in: {OUT}")


if __name__ == "__main__":
    main()
