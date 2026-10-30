"""Запуск приложения по ярлыку.

Поднимает локальный сервер на 127.0.0.1, открывает браузер и держит процесс
до остановки. Второй экземпляр не запускается: вместо этого открывается
уже работающее приложение.
"""

from __future__ import annotations

import argparse
import json
import os
import secrets
import socket
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

from . import config
from .web import create_app, START_TOKEN_PARAM

STATE_FILE = "launcher.json"


def _state_path() -> Path:
    return config.data_dir() / STATE_FILE


def _port_free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def _pick_port(preferred: int = 8765) -> int:
    if _port_free(preferred):
        return preferred
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _read_state() -> dict | None:
    path = _state_path()
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _running_instance() -> dict | None:
    """Проверить, действительно ли записанный экземпляр отвечает."""
    state = _read_state()
    if not state:
        return None
    url = state.get("health_url")
    if not url:
        return None
    try:
        with urllib.request.urlopen(url, timeout=1.5) as response:
            if response.status == 200:
                return state
    except Exception:
        return None
    return None


def _open_browser(port: int, token: str) -> None:
    url = f"http://127.0.0.1:{port}/?{START_TOKEN_PARAM}={token}"
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=config.APP_TITLE)
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--data-dir", type=Path, default=None)
    args = parser.parse_args(argv)

    if args.data_dir:
        os.environ["OTR_DATA_DIR"] = str(args.data_dir)
        config.data_dir  # кэширования нет, значение будет прочитано заново

    running = _running_instance()
    if running:
        url = f"http://127.0.0.1:{running['port']}/?{START_TOKEN_PARAM}={running['token']}"
        print(f"Приложение уже запущено: {url}")
        if not args.no_browser:
            webbrowser.open(url)
        return 0

    base = Path(os.environ.get("OTR_DATA_DIR") or config.data_dir())
    base.mkdir(parents=True, exist_ok=True)
    (base / "organizations").mkdir(exist_ok=True)
    (base / "settings").mkdir(exist_ok=True)

    port = _pick_port(args.port)
    token = secrets.token_urlsafe(24)
    app = create_app(data_dir=base, start_token=token)

    state = {
        "pid": os.getpid(),
        "port": port,
        "token": token,
        "health_url": f"http://127.0.0.1:{port}/healthz",
    }
    _state_path().write_text(json.dumps(state), encoding="utf-8")

    @app.get("/healthz")
    def healthz():
        return {"status": "ok", "version": config.VERSION}

    print(f"{config.APP_TITLE}. Открывается http://127.0.0.1:{port}/")
    print("Закройте это окно или нажмите Ctrl+C, чтобы остановить приложение.")
    if not args.no_browser:
        _open_browser(port, token)

    try:
        app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)
    except KeyboardInterrupt:
        pass
    finally:
        try:
            _state_path().unlink(missing_ok=True)
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
