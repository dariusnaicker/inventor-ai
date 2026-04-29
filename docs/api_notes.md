# Inventor MCP — API Notes and Known Limitations

## ProgID Gotcha

Autodesk Inventor does **not** expose version-suffixed ProgIDs.
`"Inventor.Application"` is the only registered ProgID regardless of version.
On a machine with both Inventor 2022 and 2026 installed, whichever was
registered last (typically the newest install) is the one `Dispatch` connects to.
There is no COM-level way to force `"Inventor.Application.26"` — the brief's
mention of this was aspirational, not an actual Autodesk API feature.

If you need to target a specific version, launch that version's executable
manually and then use `GetActiveObject("Inventor.Application")`.

## Unit Conversion

Inventor's COM API uses **centimetres** as its internal length unit.
Every method in `inventor_api.py` accepts millimetres and converts before
any COM call using `cm = mm / 10.0`.

Mass properties are returned as:
- Volume: cm³ → multiplied by 1000 to give mm³
- Area: cm² → multiplied by 100 to give mm²
- Centre of mass: cm coordinates → multiplied by 10 for mm

Mass itself is returned in kg and is not converted.

For Rocky DEM import: either set the STL import unit to "cm" in Rocky,
or rely on the `ExportUnits=4` (mm) option that `export_stl` sets via
the NameValueMap. Rocky's STL importer honours this option if present.

## STL Translator: Runtime DisplayName Lookup

`export_stl` locates the STL translator by iterating `ApplicationAddIns`
and matching `DisplayName.lower().contains("stl")`. This is intentionally
robust — the CLSID approach is fragile because:

- The user's brief cited `{533E9A98-FC3B-11D4-8E7E-0010B541CD80}` as the STL
  GUID. The `...CAA8` variant in the original request is the **DWF translator**
  CLSID — a known source of confusion in Inventor COM automation.
- The DisplayName approach survives Inventor reinstalls and service packs.

A fallback to `{533E9A98-FC3B-11D4-8E7E-0010B541CD80}` is included in case
the add-in's DisplayName doesn't match (e.g. localized installs).

STL resolution values passed via `NameValueMap.Add("Resolution", n)`:
- `0` = High (fine)
- `1` = Medium
- `3` = Low (coarse)

## makepy: One-Time Setup for Constants

Run once to regenerate the win32com type library stubs and enable constant
access via `win32com.client.constants`:

```
python -m win32com.client.makepy "Autodesk Inventor Object Library"
```

Without this, `inventor_api.py` falls back to raw integer literals
(documented at the top of the file) — the code works either way.

## Verifying from a Python REPL Before Running the Server

```python
import win32com.client

# 1. Check if Inventor is reachable
app = win32com.client.GetActiveObject("Inventor.Application")
print(app.Version)   # should print e.g. "26.0"

# 2. List add-ins to confirm STL translator is present
for i in range(1, app.ApplicationAddIns.Count + 1):
    addin = app.ApplicationAddIns.Item(i)
    try:
        print(addin.DisplayName)
    except Exception:
        pass
```

If `GetActiveObject` raises `pywintypes.com_error`, Inventor is not running —
launch it first, then retry.

## Known Limitations (v1)

1. **No assembly support**: Only `PartDocument` is created. Assembly
   (.iam) automation is out of scope for v1.

2. **Naive profile selection**: `extrude()` calls `sketch.Profiles.AddForSolid()`
   which selects the first (outermost) closed profile. For sketches with
   multiple closed loops (e.g. annular profiles), use the Inventor GUI
   or extend `extrude()` to accept a profile index.

3. **No cross-parameter expressions**: `add_parameter` / `set_parameter`
   accept numeric values only. Expressions referencing other parameters
   (e.g. `"Height / 2 mm"`) are not validated and may cause COM errors.

4. **Single active sketch**: `new_sketch()` overwrites the stored sketch
   reference. To work with multiple sketches, call `new_sketch()` again
   before each drawing operation.

5. **COM threading**: The COM apartment is implicit STA (single-threaded).
   Do not call `InventorAPI` methods from multiple threads without
   initializing COM with `pythoncom.CoInitialize()` per thread.
