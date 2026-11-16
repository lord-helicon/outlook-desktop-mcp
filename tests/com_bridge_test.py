"""
Unit test for OutlookBridge reconnect logic (com_bridge.py).

Pure unit test: fake `pythoncom` and `win32com.client` modules are injected
into sys.modules, so it runs on macOS/Linux without Outlook or pywin32.

Run: python tests/com_bridge_test.py
"""
import asyncio
import os
import sys
import threading
import types

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))


class FakeComError(Exception):
    pass


class FakeNamespace:
    def __init__(self, gen):
        self.gen = gen
        self.DefaultStore = types.SimpleNamespace(DisplayName="store")
        self.CurrentUser = types.SimpleNamespace(Name="user")


class FakeOutlook:
    def __init__(self, gen):
        self.gen = gen

    def GetNamespace(self, name):
        assert name == "MAPI"
        return FakeNamespace(self.gen)


dispatch_count = 0


def fake_dispatch(prog_id):
    global dispatch_count
    assert prog_id == "Outlook.Application"
    dispatch_count += 1
    return FakeOutlook(dispatch_count)


pythoncom = types.ModuleType("pythoncom")
pythoncom.com_error = FakeComError
pythoncom.CoInitialize = lambda: None
pythoncom.CoUninitialize = lambda: None
win32com = types.ModuleType("win32com")
win32com_client = types.ModuleType("win32com.client")
win32com_client.Dispatch = fake_dispatch
win32com.client = win32com_client
sys.modules["pythoncom"] = pythoncom
sys.modules["win32com"] = win32com
sys.modules["win32com.client"] = win32com_client

from outlook_desktop_mcp import com_bridge  # noqa: E402

passed = failed = 0


def check(name, cond):
    global passed, failed
    if cond:
        passed += 1
        print(f"PASS {name}")
    else:
        failed += 1
        print(f"FAIL {name}")


def new_bridge():
    global dispatch_count
    dispatch_count = 0
    b = com_bridge.OutlookBridge()
    b.start()
    return b


async def main():
    # 1. Normal call
    b = new_bridge()
    r = await b.call(lambda o, n, x: (n.gen, x * 2), 21)
    check("normal call returns value on first connection", r == (1, 42))
    check("normal call dispatches once", dispatch_count == 1)
    b.stop()

    # 2. Disconnect: reconnect, do NOT re-run the call, explain to the caller
    b = new_bridge()
    calls = []

    def sends_then_fails(o, n):
        calls.append(n.gen)
        # e.g. mail.Send() went through, then reading a property failed
        raise FakeComError(-2147417848, "disconnected", None, None)  # 0x80010108

    try:
        await b.call(sends_then_fails)
        check("disconnect raises OutlookReconnectedError", False)
    except com_bridge.OutlookReconnectedError as e:
        check("disconnect raises OutlookReconnectedError", True)
        check("message says connection restored, not retried",
              "restored" in str(e) and "not retried" in str(e))
    check("failed call ran exactly once (no duplicate side effects)", calls == [1])
    check("re-dispatched exactly once", dispatch_count == 2)
    r = await b.call(lambda o, n: n.gen)
    check("next call uses the new connection", r == 2)
    b.stop()

    # 2b. HRESULT carried in excepinfo scode also triggers reconnect
    b = new_bridge()

    def scode_fail(o, ns):
        raise FakeComError(-2147352567, "ex", (0, None, None, None, 0, 0x800706BA), None)

    try:
        await b.call(scode_fail)
        check("scode in excepinfo triggers reconnect", False)
    except com_bridge.OutlookReconnectedError:
        check("scode in excepinfo triggers reconnect", dispatch_count == 2)
    b.stop()

    # 2c. "Call may have executed" codes reconnect too (safe: no re-run)
    b = new_bridge()
    calls = []

    def died(o, ns):
        calls.append(1)
        raise FakeComError(0x80010007, "server died", None, None)  # RPC_E_SERVER_DIED

    try:
        await b.call(died)
    except com_bridge.OutlookReconnectedError:
        pass
    check("RPC_E_SERVER_DIED reconnects without re-running", calls == [1] and dispatch_count == 2)
    b.stop()

    # 2d. Reconnect itself fails -> ConnectionError with a clear message
    b = new_bridge()
    original = win32com_client.Dispatch

    def dispatch_down(prog_id):
        raise FakeComError(0x800401E3, "Operation unavailable", None, None)

    def gone(o, ns):
        win32com_client.Dispatch = dispatch_down
        raise FakeComError(0x800706BA, "gone", None, None)

    try:
        await b.call(gone)
        check("failed reconnect raises ConnectionError", False)
    except com_bridge.OutlookReconnectedError:
        check("failed reconnect raises ConnectionError", False)
    except ConnectionError as e:
        check("failed reconnect raises ConnectionError", "could not reconnect" in str(e))
    finally:
        win32com_client.Dispatch = original
    b.stop()

    # 3. No retry for other errors
    b = new_bridge()
    calls = []

    def other(o, ns):
        calls.append(1)
        raise FakeComError(-2147352567, "other", None, None)  # DISP_E_EXCEPTION

    try:
        await b.call(other)
        check("other com_error propagates", False)
    except FakeComError:
        check("other com_error propagates", True)
    check("no retry for other com_error", len(calls) == 1 and dispatch_count == 1)

    def valerr(o, ns):
        raise ValueError("boom")

    try:
        await b.call(valerr)
        check("non-COM error propagates", False)
    except ValueError:
        check("non-COM error propagates", True)
    check("no reconnect for non-COM error", dispatch_count == 1)
    b.stop()

    # 5. Abandoned (timed-out) queued request is skipped
    b = new_bridge()
    ran = []
    ev = threading.Event()
    b._request_queue.put((lambda o, n: ran.append(1), (), {}, ev, {"abandoned": True}))
    check("abandoned request signalled", ev.wait(timeout=3))
    check("abandoned request not executed", not ran)
    b.stop()


asyncio.run(main())
print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
