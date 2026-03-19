#!/usr/bin/env python3
"""
Lightweight local runtime helper for OCR/AI buttons.

What it does:
- Starts OCR API (port 8008) on demand
- Starts Ollama (port 11434) on demand
- Stops helper-owned services after inactivity

HTTP API (127.0.0.1:8011):
- GET  /health
- POST /prepare   {"mode":"ai"}
- POST /touch     {"reason":"optional"}
"""

from __future__ import annotations

import json
import os
import signal
import socket
import subprocess
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


HOST = "127.0.0.1"
PORT = int(os.environ.get("APEIRON_AI_HELPER_PORT", "8011"))
OCR_PORT = int(os.environ.get("APEIRON_OCR_PORT", "8008"))
OLLAMA_PORT = int(os.environ.get("APEIRON_OLLAMA_PORT", "11434"))
IDLE_SECONDS = int(os.environ.get("APEIRON_AI_IDLE_SECONDS", "600"))
AI_HOLD_SECONDS = int(os.environ.get("APEIRON_AI_HOLD_SECONDS", "900"))
ADOPT_EXISTING_RUNTIME = os.environ.get("APEIRON_ADOPT_EXISTING_RUNTIME", "1") != "0"
OCR_API_PATH = Path(
    os.environ.get(
        "APEIRON_OCR_API_PATH",
        r"C:\AI_WORKSPACE\05_AUTOMATIONS\python\ocr_api\app.py",
    )
)
PYTHON_BIN = os.environ.get("APEIRON_PYTHON_BIN", sys.executable or "python")
OLLAMA_BIN = os.environ.get("APEIRON_OLLAMA_BIN", "ollama")

CREATE_FLAGS = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

STATE_LOCK = threading.Lock()
STATE: dict[str, Any] = {
    "started_at": time.time(),
    "last_activity": time.time(),
    "hold_until": 0.0,
    "ocr_proc": None,
    "ocr_started_by_helper": False,
    "ollama_proc": None,
    "ollama_started_by_helper": False,
}


def _now() -> float:
    return time.time()


def _is_port_open(port: int, host: str = HOST, timeout: float = 0.35) -> bool:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout)
    try:
        return sock.connect_ex((host, port)) == 0
    finally:
        sock.close()


def _wait_port(port: int, timeout_sec: float) -> bool:
    deadline = _now() + timeout_sec
    while _now() < deadline:
        if _is_port_open(port):
            return True
        time.sleep(0.25)
    return _is_port_open(port)


def _proc_running(proc: subprocess.Popen | None) -> bool:
  return proc is not None and proc.poll() is None


def _pids_listening_on_port(port: int) -> list[int]:
  if os.name != "nt":
    return []
  try:
    out = subprocess.check_output(["netstat", "-ano", "-p", "tcp"], text=True, encoding="utf-8", errors="ignore")
  except Exception:  # noqa: BLE001
    return []
  pids: set[int] = set()
  needle = f":{port}"
  for line in out.splitlines():
    row = line.strip()
    if not row or "LISTENING" not in row:
      continue
    if needle not in row:
      continue
    parts = row.split()
    if len(parts) < 5:
      continue
    try:
      pids.add(int(parts[-1]))
    except Exception:  # noqa: BLE001
      pass
  return sorted(pids)


def _kill_pid(pid: int) -> None:
  if pid <= 0:
    return
  if os.name == "nt":
    subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
  else:
    try:
      os.kill(pid, signal.SIGTERM)
    except Exception:  # noqa: BLE001
      pass


def _kill_process_name(image_name: str) -> None:
  if os.name != "nt":
    return
  subprocess.run(["taskkill", "/IM", image_name, "/F", "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _touch(hold_seconds: int = 0) -> None:
    with STATE_LOCK:
        STATE["last_activity"] = _now()
        if hold_seconds > 0:
            STATE["hold_until"] = max(float(STATE.get("hold_until", 0.0)), _now() + hold_seconds)


def _start_ocr_if_needed() -> tuple[bool, str]:
    if _is_port_open(OCR_PORT):
        if ADOPT_EXISTING_RUNTIME:
            with STATE_LOCK:
                STATE["ocr_started_by_helper"] = True
        return True, "already_running"
    if not OCR_API_PATH.exists():
        return False, f"ocr_api_missing:{OCR_API_PATH}"
    app_dir = str(OCR_API_PATH.parent)

    try:
        proc = subprocess.Popen(
            [
                PYTHON_BIN,
                "-m",
                "uvicorn",
                "app:app",
                "--host",
                "127.0.0.1",
                "--port",
                str(OCR_PORT),
                "--app-dir",
                app_dir,
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_FLAGS,
        )
    except Exception as exc:  # noqa: BLE001
        return False, f"ocr_start_failed:{exc}"

    ok = _wait_port(OCR_PORT, 25)
    if not ok:
        try:
            proc.terminate()
        except Exception:  # noqa: BLE001
            pass
        return False, "ocr_port_not_ready"

    with STATE_LOCK:
        STATE["ocr_proc"] = proc
        STATE["ocr_started_by_helper"] = True
    return True, "started"


def _start_ollama_if_needed() -> tuple[bool, str]:
    if _is_port_open(OLLAMA_PORT):
        if ADOPT_EXISTING_RUNTIME:
            with STATE_LOCK:
                STATE["ollama_started_by_helper"] = True
        return True, "already_running"

    try:
        proc = subprocess.Popen(
            [OLLAMA_BIN, "serve"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_FLAGS,
        )
    except Exception as exc:  # noqa: BLE001
        return False, f"ollama_start_failed:{exc}"

    ok = _wait_port(OLLAMA_PORT, 20)
    if not ok:
        try:
            proc.terminate()
        except Exception:  # noqa: BLE001
            pass
        return False, "ollama_port_not_ready"

    with STATE_LOCK:
        STATE["ollama_proc"] = proc
        STATE["ollama_started_by_helper"] = True
    return True, "started"


def _stop_proc(proc: subprocess.Popen | None) -> None:
    if proc is None or proc.poll() is not None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=3)
    except Exception:  # noqa: BLE001
        try:
            proc.kill()
        except Exception:  # noqa: BLE001
            pass


def _idle_manager() -> None:
    while True:
        time.sleep(5)
        now = _now()
        with STATE_LOCK:
            last_activity = float(STATE.get("last_activity", now))
            hold_until = float(STATE.get("hold_until", 0.0))
            deadline_ref = max(last_activity, hold_until)
            idle_for = now - deadline_ref
            should_stop = idle_for >= IDLE_SECONDS
            ocr_owned = bool(STATE.get("ocr_started_by_helper", False))
            ollama_owned = bool(STATE.get("ollama_started_by_helper", False))
            ocr_proc = STATE.get("ocr_proc")
            ollama_proc = STATE.get("ollama_proc")

        if not should_stop:
            continue

        if ocr_owned:
            if _proc_running(ocr_proc):
                _stop_proc(ocr_proc)
            else:
                for pid in _pids_listening_on_port(OCR_PORT):
                    _kill_pid(pid)
            with STATE_LOCK:
                STATE["ocr_proc"] = None
                STATE["ocr_started_by_helper"] = False
        if ollama_owned:
            if _proc_running(ollama_proc):
                _stop_proc(ollama_proc)
            else:
                _kill_process_name("ollama.exe")
            with STATE_LOCK:
                STATE["ollama_proc"] = None
                STATE["ollama_started_by_helper"] = False


class Handler(BaseHTTPRequestHandler):
    server_version = "ApeironAiHelper/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:  # noqa: D401
        # Keep helper silent to avoid noisy background logs.
        return

    def _set_headers(self, status: int = 200, content_type: str = "application/json") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.end_headers()

    def _json(self, status: int, payload: dict[str, Any]) -> None:
        self._set_headers(status=status)
        self.wfile.write(json.dumps(payload).encode("utf-8"))

    def do_OPTIONS(self) -> None:  # noqa: N802
        self._set_headers(status=204, content_type="text/plain")

    def do_GET(self) -> None:  # noqa: N802
        if self.path != "/health":
            self._json(404, {"ok": False, "error": "not_found"})
            return
        with STATE_LOCK:
            started_at = float(STATE.get("started_at", _now()))
            last_activity = float(STATE.get("last_activity", _now()))
            hold_until = float(STATE.get("hold_until", 0.0))
            idle_ref = max(last_activity, hold_until)
            idle_for = max(0.0, _now() - idle_ref)
            payload = {
                "ok": True,
                "uptime_sec": round(_now() - started_at, 1),
                "idle_for_sec": round(idle_for, 1),
                "idle_limit_sec": IDLE_SECONDS,
                "ports": {
                    "helper": PORT,
                    "ocr": OCR_PORT,
                    "ollama": OLLAMA_PORT,
                },
                "services": {
                    "ocr_api_running": _is_port_open(OCR_PORT),
                    "ollama_running": _is_port_open(OLLAMA_PORT),
                    "ocr_owned_by_helper": bool(STATE.get("ocr_started_by_helper", False)),
                    "ollama_owned_by_helper": bool(STATE.get("ollama_started_by_helper", False)),
                },
            }
        self._json(200, payload)

    def do_POST(self) -> None:  # noqa: N802
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except Exception:  # noqa: BLE001
            length = 0
        raw = self.rfile.read(length) if length > 0 else b"{}"
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except Exception:  # noqa: BLE001
            body = {}

        if self.path == "/touch":
            _touch()
            self._json(200, {"ok": True})
            return

        if self.path != "/prepare":
            self._json(404, {"ok": False, "error": "not_found"})
            return

        mode = str(body.get("mode", "ai")).strip().lower()
        started: list[str] = []
        details: dict[str, str] = {}
        errors: list[str] = []

        if mode in {"ai", "run_ai"}:
            ok_ollama, msg_ollama = _start_ollama_if_needed()
            details["ollama"] = msg_ollama
            if not ok_ollama:
                errors.append(msg_ollama)
            elif msg_ollama == "started":
                started.append("ollama")

            ok_ocr, msg_ocr = _start_ocr_if_needed()
            details["ocr_api"] = msg_ocr
            if not ok_ocr:
                errors.append(msg_ocr)
            elif msg_ocr == "started":
                started.append("ocr_api")

            _touch(hold_seconds=AI_HOLD_SECONDS)
        else:
            _touch()

        if errors:
            self._json(
                500,
                {
                    "ok": False,
                    "mode": mode,
                    "errors": errors,
                    "details": details,
                    "started": started,
                },
            )
            return

        self._json(
            200,
            {
                "ok": True,
                "mode": mode,
                "started": started,
                "details": details,
                "idle_seconds": IDLE_SECONDS,
            },
        )


def main() -> int:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    server.daemon_threads = True

    t = threading.Thread(target=_idle_manager, daemon=True)
    t.start()

    def _shutdown(signum: int, _frame: Any) -> None:  # noqa: ARG001
        with STATE_LOCK:
            ocr_proc = STATE.get("ocr_proc")
            ollama_proc = STATE.get("ollama_proc")
            ocr_owned = bool(STATE.get("ocr_started_by_helper", False))
            ollama_owned = bool(STATE.get("ollama_started_by_helper", False))
        if ocr_owned:
            _stop_proc(ocr_proc)
        if ollama_owned:
            _stop_proc(ollama_proc)
        server.shutdown()

    signal.signal(signal.SIGTERM, _shutdown)
    signal.signal(signal.SIGINT, _shutdown)

    print(f"[ai-helper] listening on http://{HOST}:{PORT} (idle {IDLE_SECONDS}s)")
    server.serve_forever(poll_interval=0.5)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
