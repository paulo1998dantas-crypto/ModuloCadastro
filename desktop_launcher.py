import os
import socket
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from pathlib import Path

import uvicorn
import typing_extensions  # noqa: F401 - ensures PyInstaller bundles typing_extensions


HOST = "127.0.0.1"
START_PORT = 8001
APP_PATH = "/cadastro/bancos"


def boot_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BOOT_DIR = boot_dir()
BOOT_LOG_DIR = BOOT_DIR / "outputs"
BOOT_LOG_PATH = BOOT_LOG_DIR / "desktop_boot.log"


def boot_log(message: str) -> None:
    try:
        BOOT_LOG_DIR.mkdir(parents=True, exist_ok=True)
        with BOOT_LOG_PATH.open("a", encoding="utf-8") as file:
            file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
    except OSError:
        return


try:
    from main import app
except Exception:
    boot_log(traceback.format_exc())
    raise


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
LOG_DIR = APP_DIR / "outputs"
LOG_PATH = LOG_DIR / "desktop.log"


def log(message: str) -> None:
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with LOG_PATH.open("a", encoding="utf-8") as file:
            file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
    except OSError:
        return


def port_is_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.2)
        return sock.connect_ex((HOST, port)) == 0


def app_is_running(port: int) -> bool:
    return port_is_open(port)


def find_port() -> int:
    port = START_PORT
    while port_is_open(port):
        if app_is_running(port):
            return port
        port += 1
    return port


def run_server(port: int) -> None:
    try:
        uvicorn.run(
            app,
            host=HOST,
            port=port,
            log_level="warning",
            access_log=False,
            http="h11",
            log_config=None,
        )
    except Exception as exc:
        log(f"Erro no servidor: {exc}")
        raise


def wait_for_server(port: int) -> None:
    for _ in range(60):
        if app_is_running(port):
            return
        time.sleep(0.5)
    raise RuntimeError("O servidor do Modulo de Cadastro nao iniciou a tempo.")


def browser_candidates():
    env_paths = [
        os.environ.get("PROGRAMFILES", ""),
        os.environ.get("PROGRAMFILES(X86)", ""),
        os.environ.get("LOCALAPPDATA", ""),
    ]
    relative_paths = [
        Path("Microsoft/Edge/Application/msedge.exe"),
        Path("Google/Chrome/Application/chrome.exe"),
    ]
    for base in env_paths:
        if not base:
            continue
        for relative in relative_paths:
            candidate = Path(base) / relative
            if candidate.exists():
                yield candidate


def open_desktop_window(url: str):
    profile_dir = APP_DIR / "browser_profile"
    for browser in browser_candidates():
        try:
            return subprocess.Popen(
                [
                    str(browser),
                    f"--app={url}",
                    f"--user-data-dir={profile_dir}",
                    "--no-first-run",
                    "--disable-extensions",
                ],
                cwd=str(APP_DIR),
            )
        except OSError as exc:
            log(f"Falha ao abrir {browser}: {exc}")
    webbrowser.open(url)
    return None


def main() -> None:
    os.chdir(APP_DIR)
    log("Abrindo Modulo de Cadastro")
    port = find_port()
    started_here = not app_is_running(port)
    log(f"Porta selecionada: {port}")

    if started_here:
        thread = threading.Thread(target=run_server, args=(port,), daemon=False)
        thread.start()
        wait_for_server(port)

    url = f"http://{HOST}:{port}{APP_PATH}"
    try:
        (LOG_DIR / "desktop_url.txt").write_text(url, encoding="utf-8")
    except OSError:
        pass

    log(f"Abrindo janela: {url}")
    process = open_desktop_window(url)
    if process is not None:
        process.wait()
        return

    while True:
        time.sleep(3600)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"Erro ao abrir aplicacao: {exc}")
        raise
