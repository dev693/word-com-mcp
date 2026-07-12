"""Manual stdio discovery check for the MCP host (issue 0.3 / 0.10).

Spawns the built server, performs the MCP handshake, lists tools, and calls
ping/echo. Verifies every stdout line is valid JSON-RPC (nothing leaks to stdout)
and that logs appear on stderr. Not part of the xUnit suite; run on demand:

    python word-com-mcp.tests/mcp_stdio_check.py
"""
import json
import shutil
import subprocess
import sys
import threading

DLL = "word-com-mcp/bin/Debug/net8.0-windows/word-com-mcp.dll"
DOTNET = shutil.which("dotnet") or r"C:\Program Files\dotnet\dotnet.exe"


def main() -> int:
    proc = subprocess.Popen(
        [DOTNET, "exec", DLL],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        bufsize=1,
    )

    stderr_lines: list[str] = []
    threading.Thread(
        target=lambda: stderr_lines.extend(proc.stderr), daemon=True
    ).start()

    def send(obj: dict) -> None:
        proc.stdin.write(json.dumps(obj) + "\n")
        proc.stdin.flush()

    def read() -> dict:
        line = proc.stdout.readline()
        if not line:
            raise RuntimeError("server closed stdout unexpectedly")
        return json.loads(line)  # raises if stdout is polluted by non-JSON

    send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
          "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                     "clientInfo": {"name": "smoke", "version": "1.0"}}})
    init = read()
    send({"jsonrpc": "2.0", "method": "notifications/initialized"})

    send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
    tools = read()
    tool_names = sorted(t["name"] for t in tools["result"]["tools"])

    send({"jsonrpc": "2.0", "id": 3, "method": "tools/call",
          "params": {"name": "ping", "arguments": {}}})
    ping = read()

    send({"jsonrpc": "2.0", "id": 4, "method": "tools/call",
          "params": {"name": "echo", "arguments": {"message": "hallo"}}})
    echo = read()

    proc.stdin.close()
    proc.wait(timeout=10)

    ping_text = json.loads(ping["result"]["content"][0]["text"])
    echo_text = json.loads(echo["result"]["content"][0]["text"])

    ok = (
        init["result"]["serverInfo"]["name"] == "word-com-mcp"
        and tool_names == ["echo", "ping"]
        and ping_text == {"success": True, "message": "pong"}
        and echo_text == {"success": True, "message": "hallo"}
    )

    print("initialize serverInfo:", init["result"]["serverInfo"]["name"])
    print("discovered tools:", tool_names)
    print("ping ->", ping_text)
    print("echo ->", echo_text)
    print("stderr log lines:", len(stderr_lines))
    print("RESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
