# Examples

Both examples require Inventor 2026 to be running. They register documents in
the active Inventor session, build geometry, and save outputs to a folder
under the repo root.

## smoke_test.py

40 method calls across 9 sections. Touches every feature wrapper at least once
and reports OK/FAIL per call so you can see at a glance which methods work
on your Inventor version.

```powershell
python examples/smoke_test.py
```

Outputs: `_smoke_test_out/`

## turbocharger.py

A simplified turbocharger cartridge: shaft + bearing housing + compressor
wheel, assembled and constrained.

```powershell
python examples/turbocharger.py
```

Outputs: `_turbo_demo_out/`

The assembly:
- housing is grounded
- shaft is concentric with housing (axis-mate on Y) and locked axially (plane-mate flush, offset -60 mm)
- wheel is concentric with housing and locked axially inside the cavity (offset +65 mm)

Drag the shaft in Inventor and only its rotation about Y will be free, exactly as it should be.

## imd254_shaft.py

A stepped centrifugal-pump shaft built end-to-end from a real 2nd-year ME drawing brief at Stellenbosch University.

```powershell
python examples/imd254_shaft.py
```

Outputs: `_imd254_shaft_out/`

What it does:
- 14-point closed polyline revolved around Y to make the stepped shaft
  (Ø12 keyway end -> Ø8 seal area -> Ø15 roller bearing seat -> Ø18 datum -> Ø15 ball bearing seat -> Ø10 f7 right end), 133 mm overall
- Offset workplane on the +X surface, sketch a rounded-end slot, cut-extrude 3 mm deep -> 15 mm end-milled keyway
- Offset workplane on the +Y end face, sketch a Ø3.3 circle, cut-extrude 10 mm deep -> M4 tapping-drill hole

What it omits (intentionally, as a clear-eyed example):
- Tolerances and geometric-dimensioning symbols (those go on the `.idw` drawing, not the part)
- Surface finish symbols
- Two-flat features on the seal section and right end
- Chamfers (0.5 x 45 deg) on every step

The original PDF brief is course material and is **not** committed here.
