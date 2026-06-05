#!/usr/bin/env python3
"""Test MCP server startup and initialize handshake."""

import json
import subprocess
import os
import time

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# Start MCP server
proc = subprocess.Popen(
    ["uv", "run", "--directory", PROJECT_DIR, "python", "-m", "office_mcp.server"],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
    bufsize=1,
)

# Send initialize request
init_req = json.dumps({
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "test-client", "version": "0.1.0"}
    }
}) + "\n"

print(f"Sending: {init_req.strip()}")

try:
    proc.stdin.write(init_req)
    proc.stdin.flush()
except BrokenPipeError as e:
    print(f"Broken pipe: {e}")
    stderr_output = proc.stderr.read()
    print(f"Stderr: {stderr_output}")
    proc.terminate()
    exit(1)

# Wait for response
time.sleep(2)

if proc.poll() is not None:
    print(f"Server exited with code: {proc.returncode}")
    stderr_output = proc.stderr.read()
    print(f"Stderr: {stderr_output}")
else:
    # Try to read response
    try:
        # Read with timeout
        response_line = proc.stdout.readline()
        print(f"Response: {response_line.strip()}")
    except Exception as e:
        print(f"Error reading response: {e}")
        pass

    # Also print stderr
    stderr_output = proc.stderr.read()
    if stderr_output:
        print(f"Stderr: {stderr_output}")

proc.terminate()
print("Test complete.")
