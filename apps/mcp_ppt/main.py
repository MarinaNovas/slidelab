from mcp.server.fastmcp import FastMCP
from src.mcp_server.tools import register_tools


# Create an MCP server
mcp = FastMCP("PowerPoint Creator", port=8001)

register_tools(mcp)
if __name__ == "__main__":
    mcp.run(transport = "streamable-http")
