"""Privacy-conscious local counters for PowerPoint bridge traffic.

Only request/response sizes and bounded operation metadata are written to disk.
Text token counts are a rough character-based estimate, not model usage.
"""

from __future__ import annotations

from collections import Counter, deque
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import threading
from typing import Any, Mapping


_MAX_LOG_BYTES = 2 * 1024 * 1024
_HISTORY_SIZE = 256
_BROAD_READ_CHARS = 12_000
_LARGE_REPLY_CHARS = 8_000
_HASH_DIGEST_CHARS = 16

# Only these operation labels can reach the log. Unknown request values become
# "other", so an arbitrary caller string is never copied into metrics.jsonl.
_KNOWN_OPS = frozenset(
    {
        "help",
        "deck",
        "context",
        "outline",
        "selection",
        "inspect",
        "check",
        "batch",
        "render",
        "render_shape",
        "snapshot",
        "save",
        "native_read",
        "native_deck",
        "native",
        "navigate",
        "fit_text",
        "xml",
        "review",
        "metrics",
    }
)

_READ_OPS = frozenset(
    {
        "help",
        "deck",
        "context",
        "outline",
        "selection",
        "inspect",
        "check",
        "native_read",
        "native_deck",
        "native",
        "xml",
        "review",
        "metrics",
    }
)
_MUTATION_OPS = frozenset({"batch", "fit_text", "save"})
_RENDER_OPS = frozenset({"render", "render_shape", "review"})
_REPEATED_READ_OPS = frozenset({"deck", "context", "outline", "inspect"})


def _default_log_path() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        local_app_data = str(Path.home() / "AppData" / "Local")
    return Path(local_app_data) / "Codex" / "powerpoint-live" / "metrics.jsonl"


def _json_bytes(value: Any) -> bytes:
    """Serialize a JSON-compatible value for size measurement only."""
    return json.dumps(
        value,
        ensure_ascii=False,
        allow_nan=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def _sizes(value: Any) -> tuple[int, int]:
    try:
        encoded = _json_bytes(value)
    except Exception:
        # Invalid synthetic inputs must not make the bridge fail. Real bridge
        # requests and responses are JSON-compatible before reaching this API.
        return 0, 0
    return len(encoded.decode("utf-8")), len(encoded)


def _token_estimate(characters: int) -> int:
    return (characters + 3) // 4


def _revision(response: Any) -> Any:
    if not isinstance(response, Mapping):
        return None
    result = response.get("result")
    if isinstance(result, Mapping) and "revision" in result:
        return result.get("revision")
    if isinstance(result, Mapping):
        inspection = result.get("inspection")
        if isinstance(inspection, Mapping) and "revision" in inspection:
            return inspection.get("revision")
    if isinstance(result, Mapping) and 'image_hash' in result:
        return result['image_hash']
    if "revision" in response:
        return response.get("revision")
    return None


def _digest(value: Any) -> str:
    try:
        encoded = _json_bytes(value)
    except Exception:
        encoded = b""
    return hashlib.sha256(encoded).hexdigest()[:_HASH_DIGEST_CHARS]


def _request_signature(request: Any, include_id: bool = False) -> Any:
    if not isinstance(request, Mapping):
        return {}
    ignored = {"request_id"}
    if not include_id:
        ignored.add("id")
    return {
        key: value
        for key, value in request.items()
        if key not in ignored
    }


def _state_value(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return value[:128]
    if isinstance(value, (list, tuple)):
        return [_state_value(item) for item in value[:32]]
    return None


def _collect_context(value: Any, context: dict[str, Any], depth: int = 0) -> None:
    if depth > 2 or not isinstance(value, Mapping):
        return
    for key in ("deck", "slide_id", "active_slide_id", "selection_type", "ids"):
        if key in value:
            state_value = _state_value(value.get(key))
            if state_value is not None:
                context[key] = state_value
    for key in ("active", "active_slide", "slide", "context"):
        nested = value.get(key)
        if isinstance(nested, Mapping):
            _collect_context(nested, context, depth + 1)


def _active_slide_context(request: Any, response: Any) -> dict[str, Any] | None:
    context: dict[str, Any] = {}
    _collect_context(request, context)
    _collect_context(response, context)
    if isinstance(response, Mapping):
        _collect_context(response.get("result"), context)
    return context or None


def _read_state_key(request: Any, response: Any, op: str) -> str | None:
    revision = _revision(response)
    context = _active_slide_context(request, response)
    if revision is None and context is None and op != "deck":
        return None
    response_state = response.get("result") if isinstance(response, Mapping) else response
    return _digest(
        {
            "op": op,
            "request": _request_signature(request),
            "revision": revision,
            "context": context,
            "response": _digest(response_state),
        }
    )


def _render_state_key(request: Any, response: Any, op: str) -> str | None:
    image_hash = None
    if isinstance(response, Mapping):
        result = response.get("result")
        if isinstance(result, Mapping):
            image_hash = result.get("image_hash")
    if image_hash is None:
        image_hash = _revision(response)
    if image_hash is None:
        return None
    signature = _request_signature(request, include_id=op == "render_shape")
    if isinstance(signature, dict):
        signature = {key: value for key, value in signature.items()
                     if key not in {"path", "output_dir"}}
    return _digest(
        {
            "op": op,
            "request": signature,
            "image": image_hash,
        }
    )


class Metrics:
    """Keep compact counters and append privacy-safe records to local JSONL.

    Args:
        log_path: Optional JSONL destination. The default is
            ``%LOCALAPPDATA%/Codex/powerpoint-live/metrics.jsonl``.

    ``record(request, response, elapsed_ms)`` accepts the bridge request and
    response objects. ``summary()`` returns compact counters and the log path.
    Logging errors are counted and never propagated to the caller.
    """

    def __init__(self, log_path: str | os.PathLike[str] | None = None):
        self._log_path = (
            Path(log_path).expanduser()
            if log_path is not None
            else _default_log_path()
        )
        self._lock = threading.RLock()
        self._totals: Counter[str] = Counter()
        self._ops: Counter[str] = Counter()
        self._read_history: dict[str, deque[str]] = {
            op: deque(maxlen=_HISTORY_SIZE) for op in _REPEATED_READ_OPS
        }
        self._render_history: deque[str] = deque(maxlen=_HISTORY_SIZE)
        try:
            self._log_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            self._totals["metric_write_errors"] += 1

    def record(self, request: Any, response: Any, elapsed_ms: Any) -> None:
        """Record one completed bridge request without retaining its contents."""
        try:
            with self._lock:
                self._record(request, response, elapsed_ms)
        except Exception:
            # This instrumentation sits on an edit path: unexpected metrics
            # problems are visible in the summary and cannot abort an edit.
            try:
                with self._lock:
                    self._totals["metric_errors"] += 1
            except Exception:
                pass

    def _record(self, request: Any, response: Any, elapsed_ms: Any) -> None:
        req_chars, req_bytes = _sizes(request)
        res_chars, res_bytes = _sizes(response)
        try:
            duration = max(0, int(elapsed_ms))
        except (TypeError, ValueError, OverflowError):
            duration = 0

        op_value = request.get("op") if isinstance(request, Mapping) else None
        op = op_value if isinstance(op_value, str) and op_value in _KNOWN_OPS else "other"
        failed = int(
            isinstance(response, Mapping)
            and (response.get("ok") is False or "error" in response)
        )
        native_write = op in {"native", "native_deck"} and isinstance(request, Mapping) and (
            request.get("write") is True or request.get("mode") == "write"
        )
        mutation = int(op in _MUTATION_OPS or native_write)
        snapshot = int(not failed and op in {"snapshot", "xml"})
        broad_read = int(
            op in _READ_OPS
            and not native_write
            and res_chars > _BROAD_READ_CHARS
        )
        large_reply = int(res_chars > _LARGE_REPLY_CHARS)

        repeated_inspect = 0
        repeated_deck = 0
        repeated_context = 0
        repeated_outline = 0
        repeated_render = 0
        if not failed:
            read_history = self._read_history.get(op)
            if read_history is not None:
                key = _read_state_key(request, response, op)
                if key is not None:
                    repeated = int(key in read_history)
                    read_history.append(key)
                    if op == "inspect":
                        repeated_inspect = repeated
                    elif op == "deck":
                        repeated_deck = repeated
                    elif op == "context":
                        repeated_context = repeated
                    elif op == "outline":
                        repeated_outline = repeated
            if op in _RENDER_OPS:
                key = _render_state_key(request, response, op)
                if key is not None:
                    repeated_render = int(key in self._render_history)
                    self._render_history.append(key)

        flags = {
            "repeated_inspect": repeated_inspect,
            "repeated_deck": repeated_deck,
            "repeated_context": repeated_context,
            "repeated_outline": repeated_outline,
            "repeated_render": repeated_render,
            "broad_read": broad_read,
            "large_reply": large_reply,
            "failed": failed,
            "mutation": mutation,
            "snapshot": snapshot,
        }
        input_tokens = _token_estimate(req_chars)
        output_tokens = _token_estimate(res_chars)
        avoidable_tokens = (
            input_tokens + output_tokens
            if any(
                (
                    repeated_inspect,
                    repeated_deck,
                    repeated_context,
                    repeated_outline,
                    repeated_render,
                )
            )
            else 0
        )

        entry = {
            "timestamp": datetime.now(timezone.utc)
            .isoformat(timespec="milliseconds")
            .replace("+00:00", "Z"),
            "op": op,
            "elapsed_ms": duration,
            "request_chars": req_chars,
            "request_utf8_bytes": req_bytes,
            "response_chars": res_chars,
            "response_utf8_bytes": res_bytes,
            "estimated_request_text_tokens": input_tokens,
            "estimated_response_text_tokens": output_tokens,
            "flags": flags,
        }

        with self._lock:
            self._totals["requests"] += 1
            self._totals["elapsed_ms_total"] += duration
            self._totals["request_chars"] += req_chars
            self._totals["request_utf8_bytes"] += req_bytes
            self._totals["response_chars"] += res_chars
            self._totals["response_utf8_bytes"] += res_bytes
            self._totals["estimated_request_text_tokens"] += input_tokens
            self._totals["estimated_response_text_tokens"] += output_tokens
            self._totals["estimated_avoidable_text_tokens"] += avoidable_tokens
            self._totals["largest_reply_chars"] = max(
                self._totals["largest_reply_chars"], res_chars
            )
            self._totals["failed_requests"] += failed
            self._totals["mutation_requests"] += mutation
            self._totals["snapshots"] += snapshot
            self._totals["repeated_inspect"] += repeated_inspect
            self._totals["repeated_deck"] += repeated_deck
            self._totals["repeated_context"] += repeated_context
            self._totals["repeated_outline"] += repeated_outline
            self._totals["repeated_render"] += repeated_render
            self._totals["broad_reads"] += broad_read
            self._totals["large_replies"] += large_reply
            self._ops[op] += 1
            self._append(entry)

    def _append(self, entry: dict[str, Any]) -> None:
        try:
            self._log_path.parent.mkdir(parents=True, exist_ok=True)
            line = json.dumps(entry, ensure_ascii=False, separators=(",", ":")) + "\n"
            line_bytes = line.encode("utf-8")
            if (
                self._log_path.exists()
                and self._log_path.stat().st_size + len(line_bytes) > _MAX_LOG_BYTES
            ):
                previous = self._log_path.with_name(self._log_path.name + ".1")
                if previous.exists():
                    previous.unlink()
                os.replace(self._log_path, previous)
            with self._log_path.open("ab") as log_file:
                log_file.write(line_bytes)
        except OSError:
            self._totals["metric_write_errors"] += 1

    def summary(self) -> dict[str, Any]:
        """Return totals and flags without exposing request or response text."""
        with self._lock:
            requests = self._totals["requests"]
            elapsed_total = self._totals["elapsed_ms_total"]
            return {
                "requests": requests,
                "operation_counts": dict(sorted(self._ops.items())),
                "elapsed_ms_total": elapsed_total,
                "elapsed_ms_average": round(elapsed_total / requests, 1) if requests else 0,
                "request_chars": self._totals["request_chars"],
                "request_utf8_bytes": self._totals["request_utf8_bytes"],
                "response_chars": self._totals["response_chars"],
                "response_utf8_bytes": self._totals["response_utf8_bytes"],
                "estimated_request_text_tokens": self._totals[
                    "estimated_request_text_tokens"
                ],
                "estimated_response_text_tokens": self._totals[
                    "estimated_response_text_tokens"
                ],
                "estimated_avoidable_text_tokens": self._totals[
                    "estimated_avoidable_text_tokens"
                ],
                "estimated_total_text_tokens": self._totals[
                    "estimated_request_text_tokens"
                ]
                + self._totals["estimated_response_text_tokens"],
                "largest_reply_chars": self._totals["largest_reply_chars"],
                "failed_requests": self._totals["failed_requests"],
                "mutation_requests": self._totals["mutation_requests"],
                "snapshots": self._totals["snapshots"],
                "flags": {
                    "repeated_inspect": self._totals["repeated_inspect"],
                    "repeated_deck": self._totals["repeated_deck"],
                    "repeated_context": self._totals["repeated_context"],
                    "repeated_outline": self._totals["repeated_outline"],
                    "repeated_render": self._totals["repeated_render"],
                    "broad_reads": self._totals["broad_reads"],
                    "large_replies": self._totals["large_replies"],
                },
                "metric_write_errors": self._totals["metric_write_errors"],
                "metric_errors": self._totals["metric_errors"],
                "log_path": str(self._log_path),
                "token_estimate_note": (
                    "ceil(serialized JSON characters / 4); approximate text-size indicator, "
                    "not actual model, billed, vision, or reasoning tokens. "
                    "Avoidable estimates sum request and response text estimates only for "
                    "successful repeated same-state reads and identical renders; failed "
                    "calls are excluded"
                ),
            }
