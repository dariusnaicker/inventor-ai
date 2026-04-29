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
