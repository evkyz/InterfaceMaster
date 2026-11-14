"""Применение настроек из ConfigParser."""
import os
import subprocess
import sys
import time
import tkinter as tk
import winreg
from datetime import datetime
from tkinter import ttk, messagebox

import logger
from core import reg_utils, win_utils
from core.settings_io import _is_imaster_icon, _serialize_icon

MOD = "settings_apply"

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

IMASTER_MARKER = 'imaster.dll'
DRIVE_ICONS = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"
RECYCLE_PATH = r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon"
CONTROL_PANEL_CLSIDS = [
    '{26EE0668-A00A-44D7-9371-BEB064C98683}',
    '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}',
    '{21EC2020-3AEA-1069-A2DD-08002B30309D}',
]
DESKTOP_ICON_PATHS = {
    'user': r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon",
    'computer': r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon",
    'network': r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon",
}


# ---- helpers ----

def _log(message, log_name="apply_settings.log"):
    if logger.is_logging():
        logger.log(message, module=MOD)
    try:
        path = os.path.join(win_utils.get_desktop_path(), log_name)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {message}\n")
    except Exception:
        pass


def _deserialize_icon(value, current_dll_path):
    if not value:
        return value
    s = str(value)
    low = s.lower()
    if low.startswith(IMASTER_MARKER):
        if not current_dll_path:
            return s
        suffix = s[len(IMASTER_MARKER):]
        return f"{current_dll_path}{suffix}"
    pos = low.find(IMASTER_MARKER)
    if pos != -1 and current_dll_path:
        suffix = s[pos + len(IMASTER_MARKER):]
        return f"{current_dll_path}{suffix}"
    return s


def _ini_has_imaster_icons(config):
    for sec in ('desktop', 'disk'):
        if sec not in config:
            continue
        for key, value in config[sec].items():
            if key.startswith('icon_') and _is_imaster_icon(value):
                return True
    return False


def _get_current_imaster_path():
    try:
        from main import DLLChecker
        try:
            status = DLLChecker.get_dll_status()
            path = status.get('recommended_path')
            if path and os.path.isfile(path):
                return path
        except Exception:
            pass
        try:
            path = DLLChecker.get_dll_path_without_prompt()
            if path and os.path.isfile(path):
                return path
        except Exception:
            pass
    except Exception:
        pass
    sys_path = r"C:\Windows\System32\imaster.dll"
    if os.path.isfile(sys_path):
        return sys_path
    return None


def _ensure_dll_path_for_ini(config, parent):
    has_imaster = _ini_has_imaster_icons(config)
    if not has_imaster:
        return None, False

    try:
        from main import DLLChecker, DLLCheckerInternal
        from main import calculate_sha1, VALID_HASHES
    except ImportError:
        return None, True

    saved = DLLChecker.get_dll_path_from_registry()
    if saved and os.path.isfile(saved):
        h = calculate_sha1(saved)
        if h and h in VALID_HASHES:
            return saved, True

    if DLLChecker.copy_offered:
        status = DLLChecker.get_dll_status()
        return status.get('recommended_path'), True

    base = sys._MEIPASS if getattr(sys, 'frozen', False) \
        else os.path.dirname(os.path.abspath(__file__))
    fresh = DLLCheckerInternal(base).check_imaster_dll()
    if not fresh.get('dll_found') or not fresh.get('hash_match'):
        return None, True

    path = DLLChecker.offer_copy_to_system32(parent, "desktop")
    return path, True


# ---- public entry ----

def apply_settings(config, parent):
    if 'str' in config:
        apply_str(config, parent)
    if 'taskbar' in config:
        apply_taskbar(config, parent)
    if 'explorer' in config:
        apply_explorer(config, parent)

    dll_path, has_imaster = _ensure_dll_path_for_ini(config, parent)
    skip_imaster_sections = has_imaster and dll_path is None

    if 'desktop' in config:
        if skip_imaster_sections:
            _log("Секция [desktop]: пропущена (imaster.dll не найдена)")
        else:
            apply_desktop(config, dll_path=dll_path)

    if 'disk' in config:
        if skip_imaster_sections:
            _log("Секция [disk]: пропущена (imaster.dll не найдена)")
        else:
            apply_disk(config, dll_path=dll_path)

    if 'contextmenu' in config:
        apply_contextmenu(config, parent)
    if 'createnew' in config:
        apply_createnew(config)
    if 'sendto' in config:
        apply_sendto(config)

    _log("=" * 60)
    _log("ПРИМЕНЕНИЕ НАСТРОЕК ЗАВЕРШЕНО")
    _log("=" * 60)


# ---- секции (перенос из about.py) ----

def apply_str(config, parent):
    _log("=== ПРИМЕНЕНИЕ НАСТРОЕК ЯРЛЫКОВ ===")
    try:
        from str import Module as StrModule
        temp = ttk.Frame(parent)
        s = StrModule(temp)
        temp.pack_forget()
        changed = False
        if 'arrow_style' in config['str']:
            style = config['str']['arrow_style']
            _log(f"  arrow_style = {style}")
            mapping = {'default': 1, 'small': 2, 'none': 3}
            if style in mapping:
                s.arrow_choice.set(mapping[style])
                changed = True
        if 'shortcut_text' in config['str']:
            txt = config['str']['shortcut_text']
            _log(f"  shortcut_text = {txt}")
            if txt in ('default', 'removed'):
                s.shortcut_text_choice.set(1 if txt == 'default' else 2)
                changed = True
        if changed:
            s.apply_changes()
            _log("    Изменения применены")
        s.hide()
        temp.destroy()
    except Exception as e:
        logger.log(f"ERROR apply_str: {e}", module=MOD)
        _log(f"  ОШИБКА str: {e}")


# apply_taskbar / apply_explorer / apply_desktop / apply_disk /
# apply_contextmenu / apply_createnew / apply_sendto — переносятся
# из about.py с заменой self._log → _log, self.parent → parent,
# self._is_admin() → win_utils.is_admin(), self._get_current_imaster_path()
# → _get_current_imaster_path(), self._deserialize_icon → _deserialize_icon.