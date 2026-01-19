import os
import socket
import subprocess
import sys
import time


WATCH_DIR = r"e:\QC-攻关小组\正在进行项目\自主超链接\Autonomous-hyperlink\测试文件"
AUTO_HYPERLINK_SCRIPT = os.path.join(WATCH_DIR, "auto_hyperlink.py")

POLL_SECONDS = 2
STOP_GRACE_SECONDS = 6
LOCK_PORT = 52349


def _norm(p):
    try:
        return os.path.normcase(os.path.normpath(p))
    except Exception:
        return p


def _list_open_explorer_paths():
    ps = (
        "$ws=(New-Object -ComObject Shell.Application).Windows();"
        " $paths=@();"
        " foreach($w in $ws){"
        "  try{ $x=$w.Document.Folder.Self.Path; if($x){$paths+=$x} } catch{}"
        " };"
        " $paths | Select-Object -Unique"
    )
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=8,
        )
    except Exception:
        return []
    if r.returncode != 0:
        return []
    return [line.strip() for line in (r.stdout or "").splitlines() if line.strip()]


def _is_watch_dir_open():
    want = _norm(WATCH_DIR)
    want_prefix = want.rstrip("\\") + "\\"
    for p in _list_open_explorer_paths():
        np = _norm(p)
        if np == want or np.startswith(want_prefix):
            return True
    return False


def _start_child():
    python_exe = sys.executable
    base = os.path.dirname(python_exe)
    pythonw = os.path.join(base, "pythonw.exe")
    exe = pythonw if os.path.exists(pythonw) else python_exe
    creationflags = 0
    if os.name == "nt" and exe.lower().endswith("python.exe"):
        creationflags = subprocess.CREATE_NO_WINDOW
    return subprocess.Popen(
        [exe, AUTO_HYPERLINK_SCRIPT],
        cwd=WATCH_DIR,
        creationflags=creationflags,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _stop_child(proc):
    if not proc:
        return
    if proc.poll() is not None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=5)
        return
    except Exception:
        pass
    try:
        proc.kill()
    except Exception:
        pass


def _acquire_lock():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    s.bind(("127.0.0.1", LOCK_PORT))
    s.listen(1)
    return s


def main():
    if not os.path.isdir(WATCH_DIR):
        raise SystemExit(f"目录不存在: {WATCH_DIR}")
    if not os.path.isfile(AUTO_HYPERLINK_SCRIPT):
        raise SystemExit(f"找不到脚本: {AUTO_HYPERLINK_SCRIPT}")

    try:
        lock = _acquire_lock()
    except Exception:
        return

    child = None
    last_open_seen = 0.0

    try:
        while True:
            is_open = _is_watch_dir_open()
            now = time.time()
            if is_open:
                last_open_seen = now
                if child is None or child.poll() is not None:
                    child = _start_child()
            else:
                if child is not None and child.poll() is None:
                    if now - last_open_seen >= STOP_GRACE_SECONDS:
                        _stop_child(child)
                        child = None
            time.sleep(POLL_SECONDS)
    finally:
        try:
            _stop_child(child)
        finally:
            try:
                lock.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()

