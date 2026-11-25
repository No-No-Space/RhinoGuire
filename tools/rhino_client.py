"""
Helper client to interact with McNeel's rhino-mcp-router over stdio JSON-RPC.
"""
import sys
import os
import json
import subprocess
import argparse

ROUTER_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "rhino-mcp-router-nobrk.exe"))
if not os.path.exists(ROUTER_PATH):
    ROUTER_PATH = r"C:\Users\85aaj\AppData\Roaming\McNeel\Rhinoceros\ai\bin\rhino-mcp-router.exe"

class RhinoMcpClient:
    def __init__(self, router_path=ROUTER_PATH):
        self.proc = subprocess.Popen(
            [router_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1
        )
        self.req_id = 1
        self._init()

    def _init(self):
        req = {
            "jsonrpc": "2.0",
            "id": self._next_id(),
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "rhinoguire-agent", "version": "1.0"}
            }
        }
        res = self._call_raw(req)
        # notifications/initialized
        notif = {
            "jsonrpc": "2.0",
            "method": "notifications/initialized"
        }
        self.proc.stdin.write(json.dumps(notif) + "\n")
        self.proc.stdin.flush()
        return res

    def _next_id(self):
        curr = self.req_id
        self.req_id += 1
        return curr

    def _call_raw(self, req):
        self.proc.stdin.write(json.dumps(req) + "\n")
        self.proc.stdin.flush()
        line = self.proc.stdout.readline()
        if not line:
            err = self.proc.stderr.read()
            raise RuntimeError(f"Server closed connection. Stderr: {err}")
        return json.loads(line)

    def call_tool(self, tool_name, arguments=None):
        req = {
            "jsonrpc": "2.0",
            "id": self._next_id(),
            "method": "tools/call",
            "params": {
                "name": tool_name,
                "arguments": arguments or {}
            }
        }
        resp = self._call_raw(req)
        if "error" in resp:
            raise RuntimeError(f"Tool call error: {resp['error']}")
        return resp.get("result", {})

    def spawn_slot(self, version="8"):
        res = self.call_tool("spawn_slot", {"version": version})
        return res

    def list_slots(self):
        res = self.call_tool("list_slots", {})
        return res

    def close_slot(self, slot_id):
        return self.call_tool("close_slot", {"slot": slot_id})

    def run_csharp(self, script, slot=None):
        args = {"script": script}
        if slot:
            args["slot"] = slot
        return self.call_tool("run_csharp", args)

    def run_python(self, script, slot=None):
        args = {"script": script}
        if slot:
            args["slot"] = slot
        return self.call_tool("run_python", args)

    def run_command(self, command, slot=None):
        args = {"command": command}
        if slot:
            args["slot"] = slot
        return self.call_tool("run_command", args)

    def close(self):
        try:
            self.proc.terminate()
            self.proc.wait(timeout=2)
        except Exception:
            self.proc.kill()


def main():
    parser = argparse.ArgumentParser(description="Call Rhino MCP tools")
    parser.add_argument("cmd", choices=["list_slots", "spawn", "close", "csharp", "python", "cmd", "call"])
    parser.add_argument("--slot", help="Slot ID", default=None)
    parser.add_argument("--version", help="Rhino version", default="8")
    parser.add_argument("--script", help="Script text or file", default="")
    parser.add_argument("--tool", help="Tool name if cmd is call", default="")
    parser.add_argument("--json", help="Raw JSON args", default="{}")

    args = parser.parse_args()

    client = RhinoMcpClient()
    try:
        if args.cmd == "list_slots":
            res = client.list_slots()
        elif args.cmd == "spawn":
            res = client.spawn_slot(args.version)
        elif args.cmd == "close":
            res = client.close_slot(args.slot)
        elif args.cmd == "csharp":
            script = args.script
            if script.endswith(".cs") and os.path.exists(script):
                with open(script, "r", encoding="utf-8") as f:
                    script = f.read()
            res = client.run_csharp(script, args.slot)
        elif args.cmd == "python":
            script = args.script
            if script.endswith(".py") and os.path.exists(script):
                with open(script, "r", encoding="utf-8") as f:
                    script = f.read()
            res = client.run_python(script, args.slot)
        elif args.cmd == "cmd":
            res = client.run_command(args.script, args.slot)
        elif args.cmd == "call":
            params = json.loads(args.json)
            res = client.call_tool(args.tool, params)
        print(json.dumps(res, indent=2))
    finally:
        client.close()

if __name__ == "__main__":
    main()
