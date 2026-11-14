"""Определение версии/редакции Windows, права администратора,
стандартные пути, перезапуск Explorer."""
import os
import platform
import subprocess
import sys
import time
import winreg

import logger

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

_VERSION_CACHE = None
_EDITION_CACHE = None
_IS_ADMIN_CACHE = None


def _run_reg(args, **kwargs):
    kwargs.setdefault("creationflags", CREATE_NO_WINDOW)
    kwargs.setdefault("capture_output", True)
    kwargs.setdefault("text", True)
    kwargs.setdefault("encoding", "cp866")
    return subprocess.run(['reg', *args], **kwargs)


def get_windows_version():
    """Полная версия Windows: major.minor.build.revision.

    Источники (по приоритету): wmic → CurrentBuildNumber+UBR → sys.getwindowsversion.
    Результат кэшируется.
    """
    global _VERSION_CACHE
    if _VERSION_CACHE is not None:
        return _VERSION_CACHE

    try:
        r = _run_reg(['query',
                      r'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion',
                      '/v', 'CurrentBuildNumber'])
        import re
        m = re.search(r'CurrentBuildNumber\s+REG_SZ\s+(\d+)', r.stdout or '')
        if m:
            build = m.group(1)
            rev = 0
            try:
                r2 = _run_reg(['query',
                               r'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion',
                               '/v', 'UBR'])
                ubr = re.search(r'UBR\s+REG_DWORD\s+0x([0-9a-fA-F]+)', r2.stdout or '')
                if ubr:
                    rev = int(ubr.group(1), 16)
            except Exception:
                pass
            _VERSION_CACHE = f"10.0.{build}.{rev}"
            return _VERSION_CACHE
    except Exception:
        pass

    try:
        v = sys.getwindowsversion()
        _VERSION_CACHE = f"{v.major}.{v.minor}.{v.build}.{v.service_pack_major or 0}"
        return _VERSION_CACHE
    except Exception:
        pass

    try:
        parts = platform.version().split('.')
        if len(parts) >= 3:
            _VERSION_CACHE = platform.version()
            return _VERSION_CACHE
    except Exception:
        pass

    _VERSION_CACHE = "10.0.0.0"
    return _VERSION_CACHE


def get_build_number():
    try:
        return int(get_windows_version().split('.')[2])
    except Exception:
        return 0


def get_revision_number():
    try:
        parts = get_windows_version().split('.')
        return int(parts[3]) if len(parts) > 3 else 0
    except Exception:
        return 0


def get_edition_id():
    """EditionID из реестра: Professional, Enterprise, EnterpriseS, ..."""
    global _EDITION_CACHE
    if _EDITION_CACHE is not None:
        return _EDITION_CACHE
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        ) as key:
            value, _ = winreg.QueryValueEx(key, 'EditionID')
            _EDITION_CACHE = value if isinstance(value, str) else ''
    except (FileNotFoundError, OSError, TypeError):
        _EDITION_CACHE = ''
    return _EDITION_CACHE


def is_ltsc_edition():
    return get_edition_id() == 'EnterpriseS'


def is_corporate_edition():
    return get_edition_id() in ('Enterprise', 'EnterpriseS')


def is_admin():
    global _IS_ADMIN_CACHE
    if _IS_ADMIN_CACHE is not None:
        return _IS_ADMIN_CACHE
    try:
        import ctypes
        _IS_ADMIN_CACHE = bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        _IS_ADMIN_CACHE = False
    return _IS_ADMIN_CACHE


def is_windows_11():
    return get_build_number() >= 22000


def get_desktop_path():
    """Desktop через SHGetFolderPath (учитывает OneDrive/перенаправление)."""
    try:
        import ctypes
        from ctypes import wintypes
        CSIDL_DESKTOPDIRECTORY = 0x0010
        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(
            None, CSIDL_DESKTOPDIRECTORY, None, 0, buf)
        if buf.value and os.path.isdir(buf.value):
            return buf.value
    except Exception:
        pass
    p = os.path.join(os.path.expanduser("~"), "Desktop")
    return p if os.path.isdir(p) else os.path.expanduser("~")


def get_app_base_path():
    """Папка exe (frozen) или папка main.py."""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS if hasattr(sys, "_MEIPASS") else os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def restart_explorer(log_module="app"):
    """Перезапустить Проводник. Возвращает True/False."""
    try:
        subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                       capture_output=True, creationflags=CREATE_NO_WINDOW)
        time.sleep(0.5)
        subprocess.Popen(['explorer.exe'], creationflags=CREATE_NO_WINDOW)
        logger.log("explorer restarted", module=log_module)
        return True
    except Exception as e:
        logger.log(f"ERROR restart_explorer: {e}", module=log_module)
        return False


def is_paint3d_supported():
    if logger.is_debug():
        return True
    if is_ltsc_edition():
        return False
    return get_build_number() >= 15063


def is_photos_supported():
    if logger.is_debug():
        return True
    if is_ltsc_edition():
        return False
    return get_build_number() >= 15063


def is_create_video_supported():
    if logger.is_debug():
        return True
    if is_ltsc_edition():
        return False
    return get_build_number() >= 17763


def is_people_bar_supported():
    if logger.is_debug():
        return True
    if is_corporate_edition():
        return False
    return 16299 <= get_build_number() < 18363


def is_meet_now_supported():
    if logger.is_debug():
        return True
    if is_corporate_edition():
        return False
    build = get_build_number()
    rev = get_revision_number()
    if build == 19041:
        return rev >= 610
    if build == 19042:
        return rev >= 610
    return build >= 19044