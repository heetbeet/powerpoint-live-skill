"""Persistent, per-user PowerPoint COM bridge transport for Windows."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Any


_SCRIPT_PATH = Path(__file__).resolve()
_MAX_FRAME = 8 * 1024 * 1024
_DEFAULT_TIMEOUT = 30.0
_START_TIMEOUT = 12.0
_SERVER_FRAME_TIMEOUT = 60.0
_SERVER_WRITE_TIMEOUT = 5.0


class _TransportFailure(Exception):
    def __init__(self, code: str, message: str, outcome_unknown: bool = False):
        super().__init__(message)
        self.code = code
        self.outcome_unknown = outcome_unknown


def _win32() -> dict[str, Any]:
    """Load pywin32 only when a Windows transport operation is requested."""
    import pywintypes
    import win32api
    import win32con
    import win32event
    import win32file
    import win32pipe
    import win32security

    return {
        "pywintypes": pywintypes,
        "win32api": win32api,
        "win32con": win32con,
        "win32event": win32event,
        "win32file": win32file,
        "win32pipe": win32pipe,
        "win32security": win32security,
    }


def _identity() -> tuple[str, int, str]:
    import ctypes
    from ctypes import wintypes

    api = _win32()
    security = api["win32security"]
    token = security.OpenProcessToken(
        api["win32api"].GetCurrentProcess(), api["win32con"].TOKEN_QUERY
    )
    try:
        sid = security.ConvertSidToStringSid(
            security.GetTokenInformation(token, security.TokenUser)[0]
        )
    finally:
        api["win32api"].CloseHandle(token)
    session_id_out = wintypes.DWORD()
    process_id_to_session_id = ctypes.WinDLL(
        "kernel32", use_last_error=True
    ).ProcessIdToSessionId
    process_id_to_session_id.argtypes = (
        wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
    )
    process_id_to_session_id.restype = wintypes.BOOL
    if not process_id_to_session_id(os.getpid(), ctypes.byref(session_id_out)):
        raise ctypes.WinError(ctypes.get_last_error())
    session_id = session_id_out.value
    canonical = os.path.normcase(os.path.realpath(_SCRIPT_PATH))
    digest = hashlib.sha256(
        f"{sid}\0{session_id}\0{canonical}".encode("utf-8")
    ).hexdigest()[:32]
    return sid, session_id, digest


def _names() -> tuple[str, str]:
    _, _, digest = _identity()
    return (
        rf"\\.\pipe\CodexPptLive-{digest}",
        rf"Local\CodexPptLiveMutex-{digest}",
    )


def _security_attributes(sid: Any) -> Any:
    api = _win32()
    security = api["win32security"]
    if isinstance(sid, str):
        sid = security.ConvertStringSidToSid(sid)
    acl = security.ACL()
    acl.AddAccessAllowedAce(
        security.ACL_REVISION, api["win32con"].GENERIC_ALL, sid
    )
    descriptor = security.SECURITY_DESCRIPTOR()
    descriptor.SetSecurityDescriptorDacl(1, acl, 0)
    attributes = api["pywintypes"].SECURITY_ATTRIBUTES()
    attributes.SECURITY_DESCRIPTOR = descriptor
    attributes.bInheritHandle = False
    return attributes


def _error(
    code: str,
    message: str,
    elapsed_ms: int = 0,
    request_id: str | int | None = None,
    details: Any = None,
    outcome_unknown: bool = False,
) -> dict[str, Any]:
    error: dict[str, Any] = {"code": code, "message": message}
    if details is not None:
        error["details"] = details
    if outcome_unknown:
        error["outcome_unknown"] = True
    response: dict[str, Any] = {
        "ok": False,
        "error": error,
        "elapsed_ms": max(0, int(elapsed_ms)),
    }
    if request_id is not None:
        response["id"] = request_id
    return response


def _success(result: Any, elapsed_ms: int = 0, request_id: str | int | None = None) -> dict[str, Any]:
    response: dict[str, Any] = {
        "ok": True,
        "result": result,
        "elapsed_ms": max(0, int(elapsed_ms)),
    }
    if request_id is not None:
        response["id"] = request_id
    return response


def _request_id(value: Any = None) -> str | int:
    if isinstance(value, str) and value and len(value) <= 128:
        return value
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    return uuid.uuid4().hex


def _json_bytes(value: Any) -> bytes:
    return json.dumps(
        value, ensure_ascii=False, separators=(",", ":"), allow_nan=False
    ).encode("utf-8")


def _winerror(exc: BaseException) -> int | None:
    value = getattr(exc, "winerror", None)
    if value is not None:
        return value
    args = getattr(exc, "args", ())
    return args[0] if args and isinstance(args[0], int) else None


def _remaining_ms(deadline: float) -> int:
    return max(0, int((deadline - time.monotonic()) * 1000))


def _new_overlapped(api: dict[str, Any]) -> Any:
    event = api["win32event"].CreateEvent(None, True, False, None)
    overlapped = api["pywintypes"].OVERLAPPED()
    overlapped.hEvent = event
    return overlapped


def _cancel_io(handle: Any, api: dict[str, Any], overlapped: Any) -> None:
    try:
        api["win32file"].CancelIo(handle)
    except Exception:
        pass
    try:
        api["win32event"].WaitForSingleObject(overlapped.hEvent, 1000)
    except Exception:
        pass


def _wait_overlapped(
    handle: Any,
    overlapped: Any,
    api: dict[str, Any],
    deadline: float | None,
    pump: Any = None,
) -> None:
    event = overlapped.hEvent
    while True:
        wait_ms = 500 if deadline is None else min(500, _remaining_ms(deadline))
        if deadline is not None and wait_ms <= 0:
            _cancel_io(handle, api, overlapped)
            raise TimeoutError("Named pipe operation timed out")
        if pump is None:
            result = api["win32event"].WaitForSingleObject(event, wait_ms)
        else:
            result = api["win32event"].MsgWaitForMultipleObjects(
                (event,), False, wait_ms, api["win32con"].QS_ALLINPUT
            )
        if result == api["win32con"].WAIT_OBJECT_0:
            return
        if pump is not None and result == api["win32con"].WAIT_OBJECT_0 + 1:
            pump.PumpWaitingMessages()
            continue
        if result != api["win32con"].WAIT_TIMEOUT:
            _cancel_io(handle, api, overlapped)
            raise OSError("Unexpected named pipe wait result")


def _read_exact(
    handle: Any,
    count: int,
    api: dict[str, Any],
    deadline: float,
    pump: Any = None,
) -> bytes:
    chunks: list[bytes] = []
    remaining = count
    while remaining:
        buffer = api["win32file"].AllocateReadBuffer(remaining)
        overlapped = _new_overlapped(api)
        try:
            try:
                status, _ = api["win32file"].ReadFile(handle, buffer, overlapped)
            except api["pywintypes"].error as exc:
                status = _winerror(exc)
                if status != 997:  # ERROR_IO_PENDING
                    raise
            if status == 997:
                _wait_overlapped(handle, overlapped, api, deadline, pump)
            elif status not in (None, 0):
                raise OSError(f"Named pipe read failed ({status})")
            transferred = api["win32file"].GetOverlappedResult(
                handle, overlapped, False
            )
            if transferred <= 0:
                raise EOFError("Named pipe closed during read")
            chunks.append(bytes(buffer[:transferred]))
            remaining -= transferred
        finally:
            api["win32file"].CloseHandle(overlapped.hEvent)
    return b"".join(chunks)


def _write_all(
    handle: Any,
    data: bytes,
    api: dict[str, Any],
    deadline: float,
    pump: Any = None,
) -> None:
    offset = 0
    while offset < len(data):
        overlapped = _new_overlapped(api)
        try:
            try:
                status, _ = api["win32file"].WriteFile(
                    handle, data[offset:], overlapped
                )
            except api["pywintypes"].error as exc:
                status = _winerror(exc)
                if status != 997:
                    raise
            if status == 997:
                _wait_overlapped(handle, overlapped, api, deadline, pump)
            elif status not in (None, 0):
                raise OSError(f"Named pipe write failed ({status})")
            transferred = api["win32file"].GetOverlappedResult(
                handle, overlapped, False
            )
            if transferred <= 0:
                raise EOFError("Named pipe closed during write")
            offset += transferred
        finally:
            api["win32file"].CloseHandle(overlapped.hEvent)


def _write_frame(
    handle: Any,
    value: Any,
    api: dict[str, Any],
    deadline: float,
    pump: Any = None,
) -> None:
    payload = _json_bytes(value)
    if len(payload) > _MAX_FRAME:
        raise ValueError("Frame exceeds the configured size limit")
    _write_all(handle, len(payload).to_bytes(4, "big") + payload, api, deadline, pump)


def _read_frame(handle: Any, api: dict[str, Any], deadline: float, pump: Any = None) -> Any:
    length = int.from_bytes(_read_exact(handle, 4, api, deadline, pump), "big")
    if length <= 0 or length > _MAX_FRAME:
        raise ValueError("Invalid named pipe frame size")
    payload = _read_exact(handle, length, api, deadline, pump)
    return json.loads(payload.decode("utf-8"))


def _open_pipe(deadline: float, api: dict[str, Any]) -> Any:
    pipe_name, _ = _names()
    win32pipe = api["win32pipe"]
    win32file = api["win32file"]
    win32con = api["win32con"]
    while True:
        remaining = _remaining_ms(deadline)
        if remaining <= 0:
            raise _TransportFailure("connect_timeout", "Timed out waiting for the bridge")
        try:
            win32pipe.WaitNamedPipe(pipe_name, min(remaining, 250))
            return win32file.CreateFile(
                pipe_name,
                win32con.GENERIC_READ | win32con.GENERIC_WRITE,
                0,
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_FLAG_OVERLAPPED,
                None,
            )
        except api["pywintypes"].error as exc:
            code = _winerror(exc)
            if code == 231:  # ERROR_PIPE_BUSY
                continue
            if code in (2, 3):
                if _mutex_exists():
                    # A serialized server briefly removes the pipe between
                    # clients. Retry only connection setup, never a sent call.
                    time.sleep(min(0.005,max(0,deadline-time.monotonic())))
                    continue
                raise _TransportFailure("server_unavailable", "Bridge server is not running") from None
            if code == 121:
                continue
            raise _TransportFailure("pipe_connect_failed", "Could not connect to bridge") from None


def _exchange(envelope: dict[str, Any], deadline: float) -> dict[str, Any]:
    api = _win32()
    handle = None
    write_started = False
    try:
        handle = _open_pipe(deadline, api)
        write_started = True
        _write_frame(handle, envelope, api, deadline)
        response = _read_frame(handle, api, deadline)
        if not isinstance(response, dict) or not isinstance(response.get("ok"), bool):
            raise _TransportFailure("invalid_response", "Bridge returned an invalid response", True)
        try:
            _write_all(handle, b'\x06', api, deadline)
        except Exception:
            pass  # The complete response already establishes the outcome.
        return response
    except _TransportFailure:
        raise
    except TimeoutError as exc:
        raise _TransportFailure(
            "request_timeout",
            "Timed out waiting for the bridge; the operation outcome is unknown",
            write_started,
        ) from exc
    except Exception as exc:
        raise _TransportFailure(
            "server_disconnected" if write_started else "pipe_error",
            "Bridge connection ended before a complete response arrived",
            write_started,
        ) from exc
    finally:
        if handle is not None:
            try:
                api["win32file"].CloseHandle(handle)
            except Exception:
                pass


def _probe(deadline: float) -> bool:
    try:
        response = _exchange({"op": "ping"}, deadline)
        return response.get("ok") is True and response.get("result", {}).get("running") is True
    except Exception:
        return False


def _mutex_exists() -> bool:
    api = _win32()
    _, mutex_name = _names()
    try:
        handle = api["win32event"].OpenMutex(
            api["win32con"].SYNCHRONIZE, False, mutex_name
        )
    except api["pywintypes"].error:
        return False
    api["win32api"].CloseHandle(handle)
    return True


def _launch_server() -> None:
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    subprocess.Popen(
        [sys.executable, "-u", str(_SCRIPT_PATH), "--serve"],
        cwd=str(_SCRIPT_PATH.parent),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
        creationflags=creation_flags,
    )


def _ensure_server(deadline: float) -> tuple[bool, str | None]:
    now = time.monotonic()
    if _probe(min(deadline, now + 0.15)):
        return False, None
    try:
        if not _mutex_exists():
            _launch_server()
    except Exception:
        return False, "Could not start the bridge server"
    while time.monotonic() < deadline:
        if _probe(min(deadline, time.monotonic() + 0.25)):
            return True, None
        time.sleep(min(0.05, max(0.0, deadline - time.monotonic())))
    return True, "Bridge server did not become ready before the startup timeout"


def call(request: dict[str, Any], timeout: float = _DEFAULT_TIMEOUT) -> dict[str, Any]:
    """Send exactly one bridge request, returning its JSON-compatible response."""
    started_at = time.monotonic()
    request_id = _request_id(request.get("id") if isinstance(request, dict) else None)
    if not isinstance(request, dict):
        return _error("invalid_request", "Request must be a JSON object", request_id=request_id)
    if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or timeout <= 0:
        return _error("invalid_timeout", "Timeout must be a positive number", request_id=request_id)
    try:
        _json_bytes(request)
    except (TypeError, ValueError, UnicodeError):
        return _error("invalid_request", "Request must contain JSON-compatible values", request_id=request_id)

    deadline = started_at + float(timeout)
    spawned, startup_error = _ensure_server(deadline)
    if startup_error:
        code = "server_start_failed" if "Could not start" in startup_error else "startup_timeout"
        return _error(
            code,
            startup_error,
            int((time.monotonic() - started_at) * 1000),
            request_id,
        )
    del spawned
    try:
        response = _exchange(
            {"op": "request", "id": request_id, "request": request}, deadline
        )
        response["id"] = request_id
        response["elapsed_ms"] = int((time.monotonic() - started_at) * 1000)
        return response
    except _TransportFailure as exc:
        return _error(
            exc.code,
            str(exc),
            int((time.monotonic() - started_at) * 1000),
            request_id,
            outcome_unknown=exc.outcome_unknown,
        )


def _server_error(exc: BaseException, request_id: str | int) -> dict[str, Any]:
    code = getattr(exc, "code", "bridge_error")
    if not isinstance(code, (str, int)):
        code = str(code)
    message = str(exc) or type(exc).__name__
    if len(message) > 2000:
        message = message[:2000]
    details = getattr(exc, "details", None)
    response = _error(str(code), message, request_id=request_id, details=details)
    try:
        _json_bytes(response)
    except (TypeError, ValueError, UnicodeError):
        response = _error(str(code), message, request_id=request_id)
    return response


def _send_server_response(
    handle: Any, response: dict[str, Any], api: dict[str, Any], pump: Any
) -> None:
    try:
        payload = _json_bytes(response)
        if len(payload) > _MAX_FRAME:
            raise ValueError("Response exceeds frame limit")
    except (TypeError, ValueError, UnicodeError):
        response = _error("response_encoding_error", "Bridge response could not be encoded")
    try:
        _write_frame(
            handle,
            response,
            api,
            time.monotonic() + _SERVER_WRITE_TIMEOUT,
            pump,
        )
        # DisconnectNamedPipe discards unread bytes. Wait for receipt, bounded so
        # a dead client cannot hold the COM dispatcher indefinitely.
        _read_exact(handle, 1, api, time.monotonic() + 2.0, pump)
    except Exception:
        pass


def _serve_client(handle: Any, bridge: Any, api: dict[str, Any], pump: Any) -> bool:
    try:
        envelope = _read_frame(
            handle,
            api,
            time.monotonic() + _SERVER_FRAME_TIMEOUT,
            pump,
        )
    except Exception as exc:
        if os.environ.get('PPT_LIVE_DEBUG'):
            print('read envelope:', repr(exc), file=sys.stderr, flush=True)
        return False
    if not isinstance(envelope, dict):
        _send_server_response(handle, _error("invalid_envelope", "Invalid request envelope"), api, pump)
        return False

    operation = envelope.get("op")
    if operation == "ping":
        _send_server_response(handle, _success({"running": True}), api, pump)
        return False
    if operation == "stop":
        request_id = _request_id(envelope.get("id"))
        _send_server_response(handle, _success({"stopping": True}, request_id=request_id), api, pump)
        return True
    if operation != "request":
        _send_server_response(handle, _error("invalid_operation", "Unknown bridge operation"), api, pump)
        return False

    request_id = _request_id(envelope.get("id"))
    request = envelope.get("request")
    if not isinstance(request, dict):
        _send_server_response(
            handle,
            _error("invalid_request", "Request must be a JSON object", request_id=request_id),
            api,
            pump,
        )
        return False
    before = time.monotonic()
    try:
        result = bridge.metrics.summary() if request.get('op') == 'metrics' else bridge.handle(request)
        if not isinstance(result, dict):
            raise TypeError("Bridge.handle must return a dictionary")
        response = _success(result, int((time.monotonic() - before) * 1000), request_id)
        try:
            encoded = _json_bytes(response)
            if len(encoded) > _MAX_FRAME:
                raise ValueError("Bridge result exceeds frame limit")
        except (TypeError, ValueError, UnicodeError):
            response = _error(
                "invalid_bridge_response",
                "Bridge returned a value that cannot be transported",
                int((time.monotonic() - before) * 1000),
                request_id,
            )
    except Exception as exc:
        response = _server_error(exc, request_id)
        response["elapsed_ms"] = int((time.monotonic() - before) * 1000)
    if request.get('op') != 'metrics':
        try:
            bridge.metrics.record(request, response, (time.monotonic()-before)*1000)
        except Exception:
            pass  # Metrics must never turn a completed edit into a failed edit.
    _send_server_response(handle, response, api, pump)
    return False


def _serve_loop(bridge: Any, api: dict[str, Any], sid: Any, pythoncom: Any) -> None:
    pipe_name, _ = _names()
    security_attributes = _security_attributes(sid)
    win32pipe = api["win32pipe"]
    win32file = api["win32file"]
    win32con = api["win32con"]
    pipe_flags = win32pipe.PIPE_ACCESS_DUPLEX | win32con.FILE_FLAG_OVERLAPPED
    pipe_mode = win32pipe.PIPE_TYPE_BYTE | win32pipe.PIPE_READMODE_BYTE | win32pipe.PIPE_WAIT

    while True:
        pipe = win32pipe.CreateNamedPipe(
            pipe_name,
            pipe_flags,
            pipe_mode,
            1,
            65536,
            65536,
            0,
            security_attributes,
        )
        stop_requested = False
        try:
            overlapped = _new_overlapped(api)
            try:
                try:
                    connected = win32pipe.ConnectNamedPipe(pipe, overlapped)
                    status = connected if isinstance(connected, int) else 0
                except api["pywintypes"].error as exc:
                    status = _winerror(exc)
                if status == 997:  # ERROR_IO_PENDING
                    _wait_overlapped(pipe, overlapped, api, None, pythoncom)
                    win32file.GetOverlappedResult(pipe, overlapped, False)
                elif status not in (0, 535):  # ERROR_PIPE_CONNECTED
                    raise OSError(f"Named pipe accept failed ({status})")
            finally:
                win32file.CloseHandle(overlapped.hEvent)
            stop_requested = _serve_client(pipe, bridge, api, pythoncom)
        except Exception as exc:
            if os.environ.get('PPT_LIVE_DEBUG'):
                print('pipe loop:', repr(exc), file=sys.stderr, flush=True)
        finally:
            try:
                win32pipe.DisconnectNamedPipe(pipe)
            except Exception:
                pass
            win32file.CloseHandle(pipe)
        if stop_requested:
            return


def _server_main() -> int:
    api = _win32()
    sid, _, _ = _identity()
    _, mutex_name = _names()
    security_attributes = _security_attributes(sid)
    mutex = api["win32event"].CreateMutex(security_attributes, True, mutex_name)
    already_exists = api["win32api"].GetLastError() == 183  # ERROR_ALREADY_EXISTS
    if already_exists:
        api["win32api"].CloseHandle(mutex)
        return 0

    pythoncom = None
    initialized = False
    bridge = None
    try:
        import pythoncom as pythoncom_module

        pythoncom = pythoncom_module
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
        from ppt_com import Bridge
        from ppt_metrics import Metrics

        bridge = Bridge()
        bridge.metrics = Metrics()
        _serve_loop(bridge, api, sid, pythoncom)
        return 0
    except Exception:
        return 1
    finally:
        bridge = None  # Drop COM references on the initializing STA thread.
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        try:
            api["win32event"].ReleaseMutex(mutex)
        except Exception:
            pass
        api["win32api"].CloseHandle(mutex)


def _cli_start(timeout: float) -> dict[str, Any]:
    started_at = time.monotonic()
    if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or timeout <= 0:
        return _error("invalid_timeout", "Timeout must be a positive number")
    started, failure = _ensure_server(started_at + float(timeout))
    if failure:
        code = "server_start_failed" if "Could not start" in failure else "startup_timeout"
        return _error(code, failure, int((time.monotonic() - started_at) * 1000))
    return _success(
        {"running": True, "started_by_call": started},
        int((time.monotonic() - started_at) * 1000),
    )


def _cli_stop(timeout: float) -> dict[str, Any]:
    started_at = time.monotonic()
    if not isinstance(timeout, (int, float)) or isinstance(timeout, bool) or timeout <= 0:
        return _error("invalid_timeout", "Timeout must be a positive number")
    if not _probe(min(started_at + float(timeout), time.monotonic() + 0.15)) and not _mutex_exists():
        return _error("server_not_running", "Bridge server is not running")
    try:
        response = _exchange(
            {"op": "stop", "id": uuid.uuid4().hex},
            started_at + float(timeout),
        )
        response["elapsed_ms"] = int((time.monotonic() - started_at) * 1000)
        return response
    except _TransportFailure as exc:
        return _error(
            exc.code,
            str(exc),
            int((time.monotonic() - started_at) * 1000),
            outcome_unknown=exc.outcome_unknown,
        )


def _cli_status() -> dict[str, Any]:
    started_at = time.monotonic()
    running = _probe(started_at + 0.25)
    if not running:
        running = _mutex_exists()
    return _success(
        {"running": running}, int((time.monotonic() - started_at) * 1000)
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Persistent local PowerPoint bridge")
    subparsers = parser.add_subparsers(dest="command", required=True)

    request_parser = subparsers.add_parser("request", help="send one JSON bridge request")
    request_parser.add_argument("json", nargs="?", help="request JSON; reads stdin if omitted")
    request_parser.add_argument("--file", dest="json_file", help="read request JSON from an absolute path")
    request_parser.add_argument("--timeout", type=float, default=_DEFAULT_TIMEOUT)

    start_parser = subparsers.add_parser("start", help="start the bridge server")
    start_parser.add_argument("--timeout", type=float, default=_START_TIMEOUT)

    stop_parser = subparsers.add_parser("stop", help="stop the bridge server")
    stop_parser.add_argument("--timeout", type=float, default=_DEFAULT_TIMEOUT)

    subparsers.add_parser("status", help="check whether the bridge server is running")
    return parser


def _load_request(args: argparse.Namespace) -> dict[str, Any]:
    if args.json is not None and args.json_file is not None:
        raise ValueError("Provide JSON or --file, not both")
    if args.json_file is not None:
        path = Path(args.json_file)
        if not path.is_absolute():
            raise ValueError("--file requires an absolute path")
        content = path.read_text(encoding="utf-8-sig")
    elif args.json is not None:
        content = args.json
    else:
        content = sys.stdin.read()
    request = json.loads(content)
    if not isinstance(request, dict):
        raise ValueError("Request JSON must be an object")
    return request


def main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments == ["--serve"]:
        return _server_main()
    parser = _parser()
    args = parser.parse_args(arguments)
    if args.command == "request":
        try:
            request = _load_request(args)
        except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
            response = _error("invalid_request", str(exc))
        else:
            response = call(request, args.timeout)
    elif args.command == "start":
        response = _cli_start(args.timeout)
    elif args.command == "stop":
        response = _cli_stop(args.timeout)
    else:
        response = _cli_status()
    sys.stdout.write(json.dumps(response, ensure_ascii=False, separators=(",", ":")))
    sys.stdout.write("\n")
    sys.stdout.flush()
    return 0 if response.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
