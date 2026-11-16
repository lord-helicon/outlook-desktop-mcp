"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments — the live COM objects that only exist on the
COM thread.

Reconnect: if a call fails with a COM "server went away" HRESULT (Outlook was
closed or restarted), the COM thread re-dispatches Outlook.Application and
retries that call once. Because the call is retried, functions passed to
bridge.call() should be safe to run twice (they fail before completing when
the server is gone, so this is normally the case).

Timeouts: call() gives up after 60s, but the COM thread is single-threaded and
cannot be interrupted. A request that timed out while still queued is skipped
(the caller already got a TimeoutError); a request already executing (e.g.
blocked on a dialog) keeps running until it returns and blocks later requests.
"""
import threading
import queue
import asyncio
import sys
import logging

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")

CALL_TIMEOUT_SECONDS = 60

# HRESULTs meaning the Outlook COM server is gone/disconnected (compared masked
# to unsigned 32 bits). Only codes where the failed call did not run (or the
# object was already dead) are listed, so a retry cannot repeat side effects
# such as sending mail. RPC_E_SERVER_DIED (0x80010007) and RPC_S_CALL_FAILED
# (0x800706BE) are deliberately excluded: "the call may have executed".
DISCONNECT_HRESULTS = frozenset({
    0x80010108,  # RPC_E_DISCONNECTED
    0x80010012,  # RPC_E_SERVER_DIED_DNE
    0x800706BA,  # RPC_S_SERVER_UNAVAILABLE
    0x800706BF,  # RPC_S_CALL_FAILED_DNE
    0x800401FD,  # CO_E_OBJNOTCONNECTED
})


def _hresults_of(exc, com_error_type):
    """Yield candidate HRESULTs (unsigned) from a pythoncom.com_error."""
    if not isinstance(exc, com_error_type) or not exc.args:
        return
    candidates = [exc.args[0]]
    # args[2] is the EXCEPINFO tuple; its last-but-one item is scode.
    try:
        candidates.append(exc.args[2][5])
    except (IndexError, TypeError):
        pass
    for c in candidates:
        if isinstance(c, int):
            yield c & 0xFFFFFFFF


class OutlookBridge:
    """Manages a dedicated COM thread for Outlook operations."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: Exception | None = None
        self._pythoncom = None
        self._win32com_client = None

    def start(self):
        """Start the COM thread. Call once at server startup."""
        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=15):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                "Outlook COM thread failed to initialize within 15s. "
                "Is Outlook Desktop (Classic) running?"
            )

    def _connect(self):
        """Dispatch Outlook.Application and the MAPI namespace (COM thread only)."""
        self._outlook = self._win32com_client.Dispatch("Outlook.Application")
        self._namespace = self._outlook.GetNamespace("MAPI")

    def _is_disconnect(self, exc) -> bool:
        return any(
            hr in DISCONNECT_HRESULTS
            for hr in _hresults_of(exc, self._pythoncom.com_error)
        )

    def _run_request(self, func, args, kwargs):
        """Run func on the COM thread; on a disconnect error, reconnect and retry once."""
        try:
            return func(self._outlook, self._namespace, *args, **kwargs)
        except Exception as e:
            if not self._is_disconnect(e):
                raise
            logger.warning("Outlook COM disconnected (%s); reconnecting", e)
            self._connect()  # if this raises, it propagates to the caller
            return func(self._outlook, self._namespace, *args, **kwargs)

    def _com_thread_main(self):
        """Main loop for the COM thread."""
        import pythoncom
        import win32com.client

        self._pythoncom = pythoncom
        self._win32com_client = win32com.client

        pythoncom.CoInitialize()
        try:
            self._connect()
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            logger.debug("COM thread ready. Store: %s, User: %s", store_name, user_name)
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                if result_holder.get("abandoned"):
                    # Caller already timed out; don't run stale work.
                    result_event.set()
                    continue
                try:
                    result_holder["value"] = self._run_request(func, args, kwargs)
                except Exception as e:
                    result_holder["error"] = e
                finally:
                    result_event.set()
        except Exception as e:
            self._init_error = e
            self._ready.set()  # Unblock the caller so they see the error
            logger.error("COM thread init failed: %s", e)
        finally:
            pythoncom.CoUninitialize()

    async def call(self, func, *args, **kwargs):
        """
        Schedule a function to run on the COM thread and await its result.

        The function signature must be: func(outlook, namespace, *args, **kwargs)
        """
        result_event = threading.Event()
        result_holder = {}
        self._request_queue.put((func, args, kwargs, result_event, result_holder))

        loop = asyncio.get_running_loop()
        signaled = await loop.run_in_executor(
            None, lambda: result_event.wait(timeout=CALL_TIMEOUT_SECONDS)
        )
        if not signaled:
            result_holder["abandoned"] = True  # COM thread skips it if still queued
            raise TimeoutError(
                "Outlook COM operation timed out after 60 seconds. "
                "Outlook may be waiting for user input (e.g., a dialog box)."
            )

        if "error" in result_holder:
            raise result_holder["error"]
        return result_holder.get("value")

    def stop(self):
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        if self._thread:
            self._thread.join(timeout=5)
