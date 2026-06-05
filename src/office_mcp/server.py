#!/usr/bin/env python3
"""Office MCP Server - Main entry point.

Combines Word, Excel, PPT tools into a single MCP server.
"""

import sys
import argparse
from fastmcp import FastMCP

# Build the combined server
mcp = FastMCP("office-mcp-server")


def main():
    parser = argparse.ArgumentParser(description="Office MCP Server")
    parser.add_argument("--transport", choices=["stdio", "sse"], default="stdio",
                        help="Transport protocol (default: stdio)")
    parser.add_argument("--port", type=int, default=8000,
                        help="Port for SSE transport (default: 8000)")
    args = parser.parse_args()

    # Import and register Word tools
    try:
        from office_mcp.word.tools import mcp as word_mcp
        for tool in word_mcp._tools.values():
            mcp.add_tool(tool)
        print("✅ Word tools registered", file=sys.stderr)
    except Exception as e:
        print(f"⚠️  Warning: Could not load Word tools: {e}", file=sys.stderr)

    # TODO: Register Excel tools (Phase 3)
    # try:
    #     from office_mcp.excel.tools import mcp as excel_mcp
    #     for tool in excel_mcp._tools.values():
    #         mcp.add_tool(tool)
    #     print("✅ Excel tools registered", file=sys.stderr)
    # except Exception as e:
    #     print(f"⚠️  Warning: Could not load Excel tools: {e}", file=sys.stderr)

    # TODO: Register PPT tools (Phase 4)
    # try:
    #     from office_mcp.ppt.tools import mcp as ppt_mcp
    #     for tool in ppt_mcp._tools.values():
    #         mcp.add_tool(tool)
    #     print("✅ PPT tools registered", file=sys.stderr)
    # except Exception as e:
    #     print(f"⚠️  Warning: Could not load PPT tools: {e}", file=sys.stderr)

    print(f"🚀 Starting Office MCP Server (transport={args.transport})", file=sys.stderr)
    
    if args.transport == "sse":
        mcp.run(transport="sse", host="127.0.0.1", port=args.port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
