"""Main entry point for Photoshop MCP server."""
from .server import mcp

if __name__ == "__main__":
    mcp.run("stdio")