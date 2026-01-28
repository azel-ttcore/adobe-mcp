"""Main entry point for Premiere MCP server."""
from .server import mcp

if __name__ == "__main__":
    mcp.run("stdio")