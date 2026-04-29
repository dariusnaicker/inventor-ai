# inventor-ai

**Drive Autodesk Inventor 2026 from Claude (or any MCP client).**

A Model Context Protocol server for Autodesk Inventor's COM API. Supports parts, assemblies, work-feature constraints, exports (STL/STEP), and a `run_python` escape hatch that gives you the full Inventor object model when a dedicated tool doesn't exist yet.

> "Build me a turbocharger cartridge in Inventor" -> three parts, one fully constrained assembly, exported STL + STEP, in one prompt.

## Why

Autodesk's first-party MCP support is for Fusion 360. Inventor users have been left out. This package fills that gap: 30+ MCP tools that cover the Inventor workflows you'd actually use day-to-day, plus an escape hatch that scales to anything else.

## Features

- **Parts**: sketches on origin planes, lines, circles, rectangles, extrude, revolve, sweep, loft, fillet, shell, mirror, circular pattern, rectangular pattern, parameters
- **Assemblies**: place components, ground, axis-mate (concentric), plane-mate (with offset, mate or flush)
- **Exports**: STL (binary or ASCII), STEP (AP203/AP214/AP242), `.ipt` / `.iam` save
- **Diagnostics**: list features, list faces, list parameters, undo, test connection
- **Escape hatch**: `run_python(code, part_name)` executes arbitrary Python with `app`, `doc`, `comp_def`, `sketch`, `tg` (transient geometry), and unit helpers pre-injected. If a feature isn't wrapped yet, you can still do it.

## Install

Requires Windows + Autodesk Inventor 2026 + Python 3.10+.

```powershell
pip install fastmcp pywin32
git clone https://github.com/<your-username>/inventor-ai.git
cd inventor-ai
pip install -e .
```

## Run the server

```powershell
python -m inventor_ai.server
```

Or register it with Claude Desktop / any MCP client. Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "inventor": {
      "command": "python",
      "args": ["-m", "inventor_ai.server"]
    }
  }
}
```

Inventor 2026 must be running before you call any tool.

## Examples

Two scripts in `examples/`:

### Smoke test (40 method calls, 9 sections)

```powershell
python examples/smoke_test.py
```

Builds a 50x30x20 box, a flange with a 6-hole bolt circle, a mirrored stub, a 3x3 rectangular post array, and round-trips a parameter. Outputs to `_smoke_test_out/`.

### Turbocharger cartridge (3 parts + constrained assembly)

```powershell
python examples/turbocharger.py
```

Builds:
1. **turbo_shaft.ipt** -- 180 mm stepped shaft (revolve + fillet)
2. **turbo_bearing_housing.ipt** -- revolved housing with a 30 mm bore (lower) and a 70 mm cavity (upper) for the wheel, plus 6 bolt holes
3. **turbo_compressor_wheel.ipt** -- hub with 26 mm bore + 8 radial blades
4. **turbocharger.iam** -- fully constrained: housing grounded, shaft and wheel axis-mated to the housing's Y axis, plane-mated for axial position. 1 free DOF (shaft rotation about Y), as it should be.

Outputs to `_turbo_demo_out/`.

## Tool list (30 MCP tools)

```
Connection      inventor_connect, inventor_test_connection, inventor_list_open_documents,
                inventor_close_document, inventor_run_python

Parts           inventor_new_part, inventor_save_part, inventor_export_stl,
                inventor_export_step

Sketching       inventor_new_sketch, inventor_draw_line, inventor_draw_rectangle,
                inventor_draw_circle

Features        inventor_extrude, inventor_revolve, inventor_sweep, inventor_loft,
                inventor_fillet_all_edges, inventor_shell, inventor_mirror,
                inventor_circular_pattern, inventor_rectangular_pattern, inventor_undo

Diagnostics     inventor_list_features, inventor_list_faces, inventor_list_parameters,
                inventor_add_parameter

Assemblies      inventor_new_assembly, inventor_place_component, inventor_save_assembly,
                inventor_ground_component, inventor_assemble_axis_mate,
                inventor_assemble_plane_mate
```

## Using `run_python` for unwrapped features

Every feature wrapper above is implemented on top of `run_python`. You can use the same escape hatch yourself:

```python
api.run_python("""
# Pre-injected: app, doc, comp_def, sketch, tg, mm_to_cm,
#               EXTENT_OP, EXTENT_DIR, PLANE_INDEX, api
edges = comp_def.SurfaceBodies.Item(1).Edges
result = {"edge_count": edges.Count}
""", part_name="my_part")
```

This is how you reach Inventor's deeper API surface (work features, sheet metal, surfaces, FEM) without waiting for a dedicated wrapper. It's also a useful debugging tool when a wrapper misbehaves.

## Security

The `run_python` / `inventor_run_python` tool executes arbitrary Python on the host with live COM access. **Never expose this server beyond localhost.** FastMCP defaults to stdio transport, which is local-only. If you ever switch to HTTP/SSE transport, bind to 127.0.0.1 only and put it behind authentication. Treat the MCP server as you would a local Jupyter kernel: never trust untrusted clients.

## Contributing

PRs welcome. Areas that would help most:
- Drawing automation (`.idw` views, dimensions)
- Sheet metal features
- iLogic rule generation
- A test harness that mocks the COM layer (so CI can run without Inventor)

When adding a new feature, prefer building the wrapper on top of `run_python` (see `fillet_all_edges`, `shell`, `loft`, etc.) rather than re-implementing COM constants -- that keeps the COM surface in one place.

## Status

Beta. Tested on Inventor 2026.2 (Build 302298010). Should work on 2025+ but COM constants are version-specific so YMMV.

## License

MIT. See `LICENSE`.
