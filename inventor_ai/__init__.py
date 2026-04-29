"""Autodesk Inventor MCP server.

A Model Context Protocol server that lets Claude (and any MCP client) drive
Autodesk Inventor 2026 through its COM API. Supports parts, assemblies,
constraints, and a `run_python` escape hatch for anything not covered by a
dedicated tool.
"""

from .api import InventorAPI

__all__ = ["InventorAPI"]
__version__ = "0.1.0"
