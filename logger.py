"""Единая точка логирования и режимов работы."""
import os
import subprocess
import sys
from datetime import datetime

LOG_MODE = False
DEBUG_MODE = False

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

_log_path = None
_init_error = None


def _console(msg):
    """Дублировать в stdout (видно в PyCharm Run / cmd)."""
    try:
        sys.stdout.write(msg + "\n")
        sys.stdout.flush()
    except Exception:
        pass


def _candidate_log_dirs():
    result = []
    # 1. Desktop через SHGetFolderPath (учитывает OneDrive/перенаправление)
    try:
        import ctypes
        from ctypes import wintypes
        CSIDL_DESKTOPDIRECTORY = 0x0010
        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(
            None, CSIDL_DESKTOPDIRECTORY, None, 0, buf)
        if buf.value and os.path.isdir(buf.value):
            result.append(buf.value)
    except Exception:
        pass
    # 2. ~/Desktop
    try:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if os.path.isdir(desktop) and desktop not in result:
            result.append(desktop)
    except Exception:
        pass
    # 3. USERPROFILE
    try:
        up = os.environ.get("USERPROFILE")
        if up and os.path.isdir(up) and up not in result:
            result.append(up)
    except Exception:
        pass
    # 4. TEMP
    try:
        tmp = os.environ.get("TEMP") or os.environ.get("TMP")
        if tmp and os.path.isdir(tmp) and tmp not in result:
            result.append(tmp)
    except Exception:
        pass
    # 5. Папка с exe / проекта
    try:
        base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
            else os.path.dirname(os.path.abspath(__file__))
        if base and os.path.isdir(base) and base not in result:
            result.append(base)
    except Exception:
        pass
    # 6. cwd
    try:
        cwd = os.getcwd()
        if cwd and os.path.isdir(cwd) and cwd not in result:
            result.append(cwd)
    except Exception:
        pass
    return result


def init_from_argv(argv):
    global LOG_MODE, DEBUG_MODE, _log_path, _init_error

    _console(f"[logger] init_from_argv argv={argv!r}")
    _console(f"[logger] frozen={getattr(sys, 'frozen', False)}")

    lower = [a.lower() for a in argv]
    LOG_MODE = any("-log" == a or a.endswith("\\-log") or a.endswith("/-log")
                   for a in lower)
    DEBUG_MODE = any("-dbg" == a or a.endswith("\\-dbg") or a.endswith("/-dbg")
                     for a in lower)
    _console(f"[logger] LOG_MODE={LOG_MODE} DEBUG_MODE={DEBUG_MODE}")

    if not (LOG_MODE or DEBUG_MODE):
        _console("[logger] no log flags, exit")
        return

    suffix = "dbg" if DEBUG_MODE else "log"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"interface_master_{suffix}_{ts}.log"

    candidates = _candidate_log_dirs()
    _console(f"[logger] candidates = {candidates!r}")

    for base in candidates:
        try:
            candidate = os.path.join(base, filename)
            with open(candidate, "a", encoding="utf-8") as f:
                f.write("")
            _log_path = candidate
            _console(f"[logger] log_path = {_log_path!r}")
            break
        except Exception as e:
            _init_error = f"{base}: {e}"
            _console(f"[logger] candidate failed: {base}: {e}")
            continue

    if _log_path:
        mode = "DEBUG" if DEBUG_MODE else "LOG"
        _write_raw(f"===== Interface Master {mode} session =====")
        _write_raw(f"argv = {argv!r}")
        _write_raw(f"cwd  = {os.getcwd()!r}")
        _write_raw(f"exe  = {sys.executable!r}")
        _console("[logger] log initialized OK")
    else:
        _console(f"[logger] cannot open any log file: {_init_error}")


def get_log_path():
    return _log_path


def _write_raw(line):
    if not _log_path:
        return
    try:
        with open(_log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        try:
            sys.stderr.write(f"[InterfaceMaster] log write error: {e}\n")
        except Exception:
            pass


def log(message, module=None):
    if not (LOG_MODE or DEBUG_MODE):
        return
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    prefix = f"[{ts}]"
    if module:
        prefix += f" [{module}]"
    line = f"{prefix} {message}"
    _write_raw(line)
    _console(line)


def log_reg(module, action, key_path, value_name=None, value_data=None):
    if not (LOG_MODE or DEBUG_MODE):
        return
    parts = [f"REG {action.upper()}", key_path]
    if value_name is not None:
        parts.append(f"\\{value_name}" if value_name else "\\<default>")
    if value_data is not None:
        parts.append(f"= {value_data!r}")
    log(" ".join(parts), module=module)


def log_file_attr(module, path, action, value=None):
    if not (LOG_MODE or DEBUG_MODE):
        return
    tail = f" -> {value!r}" if value is not None else ""
    log(f"FILE {action}: {path}{tail}", module=module)


def log_module_open(module_name):
    if DEBUG_MODE:
        log(f"MODULE OPEN: {module_name}", module="app")


def is_debug():
    return DEBUG_MODE


def is_logging():
    return LOG_MODE or DEBUG_MODE