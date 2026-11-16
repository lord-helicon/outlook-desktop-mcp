"""
Unit test for utils/output.py: tool results leave the server as compact JSON.

Pure unit test: no Outlook, runs on any platform.

Run: .venv\\Scripts\\python tests\\output_test.py
"""
import asyncio
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from mcp.server.fastmcp import FastMCP  # noqa: E402

from outlook_desktop_mcp.utils.output import compact_json, compact_json_tools  # noqa: E402


def _text(result):
    content = result[0] if isinstance(result, tuple) else result
    return content[0].text


def test_compact_json():
    pretty = json.dumps([{"subject": "Grüße", "n": 1}], indent=2)
    assert compact_json(pretty) == '[{"subject":"Grüße","n":1}]'
    assert compact_json('  {"a": [1, 2]}') == '{"a":[1,2]}'
    # Non-JSON results pass through untouched
    for s in ("Marked as read: 'x'", "Error: not found", "", "{not json", "[oops"):
        assert compact_json(s) == s, s
    assert compact_json(None) is None


def _tools(mcp):
    return {t.name: t.model_dump(mode="json") for t in asyncio.run(mcp.list_tools())}


def test_schemas_unchanged_and_output_compact():
    async def list_things(folder: str = "inbox", count: int = 10) -> str:
        """List things in a folder.

        Args:
            folder: Folder name.
            count: How many.
        """
        return json.dumps([{"folder": folder, "count": count, "name": "Jörg"}], indent=2)

    async def say(text: str) -> str:
        """Echo plain text."""
        return f"Said: {text}"

    plain = FastMCP("plain")
    plain.tool()(list_things)
    plain.tool()(say)

    compact = compact_json_tools(FastMCP("compact"))
    returned = compact.tool()(list_things)
    compact.tool()(say)

    assert _tools(plain) == _tools(compact), "tool schemas must not change"
    assert returned is list_things, "decorator returns the original function"

    out = _text(asyncio.run(compact.call_tool("list_things", {"folder": "x", "count": 3})))
    assert out == '[{"folder":"x","count":3,"name":"Jörg"}]', out
    assert _text(asyncio.run(compact.call_tool("say", {"text": "hi"}))) == "Said: hi"
    # Direct calls keep the original (indented) output
    assert asyncio.run(list_things()).startswith("[\n")


def test_sync_tool_rejected():
    mcp = compact_json_tools(FastMCP("x"))

    def not_async() -> str:
        return "x"

    try:
        mcp.tool()(not_async)
    except TypeError:
        return
    raise AssertionError("sync tools must be rejected")


if __name__ == "__main__":
    test_compact_json()
    test_schemas_unchanged_and_output_compact()
    test_sync_tool_rejected()
    print("OK")
