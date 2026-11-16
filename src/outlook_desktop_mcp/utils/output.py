"""Compact JSON for tool results.

Tool results go straight into the model's context, so whitespace costs
tokens. Tools build their JSON with ``json.dumps(..., indent=2)`` (readable
in logs and tests); ``compact_json_tools`` re-serializes it compactly at the
point the result leaves the MCP server:

- no indentation or spaces after separators (about 25-30% fewer characters
  on list results);
- ``ensure_ascii=False``, so "für" stays three characters instead of
  "f\\u00fcr".

Non-JSON results (confirmation and error strings) pass through unchanged.
Doing this once at registration, rather than at each ``json.dumps`` call,
covers every tool, including ones added later.
"""
import functools
import inspect
import json


def compact_json(result):
    """Return ``result`` re-serialized compactly if it is a JSON string."""
    if not isinstance(result, str):
        return result
    stripped = result.lstrip()
    if not stripped or stripped[0] not in "[{":
        return result
    try:
        data = json.loads(result)
    except ValueError:
        return result
    return json.dumps(data, separators=(",", ":"), ensure_ascii=False)


def compact_json_tools(mcp):
    """Make every tool later registered with ``@mcp.tool()`` return compact JSON.

    The decorator registers a wrapper with the same name, docstring and
    signature (via functools.wraps), so tool schemas are unchanged. It
    returns the original function, so direct calls (tests, other tools)
    still get the original output.
    """
    register_tool = mcp.tool

    @functools.wraps(register_tool)
    def tool(*args, **kwargs):
        register = register_tool(*args, **kwargs)

        def decorator(fn):
            if not inspect.iscoroutinefunction(fn):
                raise TypeError(f"Tool {fn.__name__} must be async")

            @functools.wraps(fn)
            async def compact_result(*call_args, **call_kwargs):
                return compact_json(await fn(*call_args, **call_kwargs))

            register(compact_result)
            return fn

        return decorator

    mcp.tool = tool
    return mcp
