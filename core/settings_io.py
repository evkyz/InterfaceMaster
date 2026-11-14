"""Сбор текущих настроек и сохранение/загрузка INI."""
import os
import re
import subprocess
import sys
import winreg
from configparser import ConfigParser
from datetime import datetime
from tkinter import filedialog, messagebox

import logger
from core import reg_utils, win_utils

MOD = "settings_io"

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

IMASTER_MARKER = 'imaster.dll'
CONTROL_PANEL_CLSIDS = [
    '{26EE0668-A00A-44D7-9371-BEB064C98683}',
    '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}',
    '{21EC2020-3AEA-1069-A2DD-08002B30309D}',
]
DRIVE_ICONS = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"
RECYCLE_INDEX_PAIRS = {
    '4': ('4', '5'), '5': ('4', '5'),
    '24': ('24', '25'), '25': ('24', '25'),
    '34': ('34', '35'), '35': ('34', '35'),
    '37': ('37', '38'), '38': ('37', '38'),
    '44': ('44', '45'), '45': ('44', '45'),
    '64': ('64', '65'), '65': ('64', '65'),
    '74': ('74', '75'), '75': ('74', '75'),
    '84': ('84', '85'), '85': ('84', '85'),
}


# ---- helpers ----

def _is_imaster_icon(value):
    return bool(value) and IMASTER_MARKER in str(value).lower()


def _serialize_icon(value):
    if not value:
        return value
    s = str(value)
    low = s.lower()
    pos = low.find(IMASTER_MARKER)
    if pos == -1:
        return s
    return s[pos:]


def _get_value(path, name='', hive='HKCU'):
    return reg_utils.get_value(hive, path, name)


def _key_exists(path, hive='HKLM'):
    return reg_utils.key_exists(hive, path)


# ---- collect ----

def collect_all_settings(version):
    return {
        'metadata': {
            'version': version,
            'export_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'program': 'Interface Master',
        },
        'desktop': collect_desktop(),
        'disk': collect_disk(),
        'str': collect_shortcut(),
        'contextmenu': collect_contextmenu(),
        'taskbar': collect_taskbar(),
        'explorer': collect_explorer(),
        'createnew': collect_createnew(),
        'sendto': collect_sendto(),
    }


def collect_desktop():
    settings = {}
    items = {
        'user': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon',
        'computer': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon',
        'network': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon',
    }
    for item, path in items.items():
        val = _get_value(path, '', 'HKCU')
        if val and _is_imaster_icon(val):
            settings[f'icon_{item}'] = _serialize_icon(val)

    recycle_reg = r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon'
    empty_default = _get_value(recycle_reg, '', 'HKCU')
    empty_explicit = _get_value(recycle_reg, 'Empty', 'HKCU')
    full_val = _get_value(recycle_reg, 'Full', 'HKCU')

    src_empty = (empty_explicit if (empty_explicit and _is_imaster_icon(empty_explicit))
                 else empty_default)
    if src_empty and _is_imaster_icon(src_empty):
        matched = False
        if ',-' in src_empty:
            idx = src_empty.split(',-')[-1]
            if idx in RECYCLE_INDEX_PAIRS:
                empty_idx, full_idx = RECYCLE_INDEX_PAIRS[idx]
                dll = src_empty.split(',-')[0]
                settings['icon_recycle_bin_empty'] = _serialize_icon(
                    f"{dll},-{empty_idx}")
                if full_val and _is_imaster_icon(full_val):
                    settings['icon_recycle_bin_full'] = _serialize_icon(full_val)
                matched = True
        if not matched:
            settings['icon_recycle_bin_empty'] = _serialize_icon(src_empty)
            if full_val and _is_imaster_icon(full_val):
                settings['icon_recycle_bin_full'] = _serialize_icon(full_val)

    libraries = _get_value(r'CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon',
                           '', 'HKCR')
    if libraries and _is_imaster_icon(libraries):
        settings['icon_libraries'] = _serialize_icon(libraries)

    for clsid in CONTROL_PANEL_CLSIDS:
        path = rf'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{clsid}\DefaultIcon'
        val = _get_value(path, '', 'HKCU')
        if val and _is_imaster_icon(val):
            settings[f'icon_control_panel_{clsid}'] = _serialize_icon(val)
    return settings


def collect_disk():
    settings = {}
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, DRIVE_ICONS) as root:
            i = 0
            while True:
                try:
                    drive = winreg.EnumKey(root, i)
                except OSError:
                    break
                i += 1
                if len(drive) != 1 or not drive.isalpha():
                    continue
                try:
                    with winreg.OpenKey(
                            winreg.HKEY_LOCAL_MACHINE,
                            f"{DRIVE_ICONS}\\{drive}\\DefaultIcon") as k:
                        val, _ = winreg.QueryValueEx(k, "")
                except FileNotFoundError:
                    continue
                if val and _is_imaster_icon(val):
                    settings[f'icon_{drive.upper()}'] = _serialize_icon(val)
    except FileNotFoundError:
        pass
    return settings


def collect_shortcut():
    settings = {}
    arrow = _get_value(
        r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons',
        '29', 'HKLM')
    if arrow is None:
        settings['arrow_style'] = 'default'
    elif arrow == '':
        settings['arrow_style'] = 'none'
    elif 'shell32.dll,-30' in str(arrow):
        settings['arrow_style'] = 'small'
    else:
        settings['arrow_style'] = 'custom'

    tmpl = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Explorer\NamingTemplates',
        'ShortcutNameTemplate', 'HKCU')
    if tmpl:
        if '"%s — ярлык.lnk"' in str(tmpl):
            settings['shortcut_text'] = 'default'
        elif '"%s.lnk"' in str(tmpl):
            settings['shortcut_text'] = 'removed'
        else:
            settings['shortcut_text'] = 'custom'
    else:
        settings['shortcut_text'] = 'default'
    return settings


def collect_contextmenu():
    settings = {}
    try:
        from contextmenu import check_menu_item_exists, CONTEXT_MENU_KEYS
        for key, item_type in CONTEXT_MENU_KEYS.items():
            try:
                settings[key] = bool(check_menu_item_exists(item_type))
            except Exception:
                settings[key] = False
    except Exception as e:
        logger.log(f"collect contextmenu error: {e}", module=MOD)
    return settings


def collect_taskbar():
    s = {}
    short_date = _get_value(r'Control Panel\International', 'sShortDate', 'HKCU')
    s['show_weekday'] = 'ddd' in str(short_date) if short_date else False

    notif = _get_value(r'SOFTWARE\Policies\Microsoft\Windows\Explorer',
                       'DisableNotificationCenter', 'HKLM')
    s['notification_center'] = notif != '0x1'

    small = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
        'TaskbarSmallIcons', 'HKCU')
    s['small_icons'] = small == '0x1'

    search = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Search',
        'SearchboxTaskbarMode', 'HKCU')
    s['search_mode'] = {'0x0': 0, '0x1': 1}.get(search, 2)

    tv = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
        'ShowTaskViewButton', 'HKCU')
    s['show_task_view'] = tv != '0x0'

    glom = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
        'TaskbarGlomLevel', 'HKCU')
    s['grouping_mode'] = {'0x0': 0, '0x1': 1, '0x2': 2}.get(glom, 0)

    try:
        if win_utils.is_people_bar_supported():
            people = _get_value(r'Software\Policies\Microsoft\Windows\Explorer',
                                'HidePeopleBar', 'HKCU')
            s['show_people'] = people != '0x1'
        if win_utils.is_meet_now_supported():
            meet = _get_value(
                r'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer',
                'HideSCAMeetNow', 'HKLM')
            s['show_meet_now'] = meet != '0x1'
    except Exception as e:
        logger.log(f"collect taskbar people/meet error: {e}", module=MOD)
    return s


def collect_explorer():
    s = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Drive\shell\encrypt-bde') as k:
            winreg.QueryValueEx(k, 'ProgrammaticAccessOnly')
            s['bitlocker_enabled'] = False
    except (FileNotFoundError, OSError):
        s['bitlocker_enabled'] = True

    photo_assoc = _get_value(
        r'SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
        '.jpg', 'HKLM')
    photo_class = _key_exists(r'SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg', 'HKLM')
    s['photo_viewer_enabled'] = (photo_assoc is not None) and photo_class

    hide_ext = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
        'HideFileExt', 'HKCU')
    s['show_file_extensions'] = hide_ext == '0x0'

    no_recent = _get_value(
        r'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer',
        'NoRecentDocsHistory', 'HKCU')
    s['recent_folders_enabled'] = no_recent != '0x1'

    try:
        from explorer import (
            SHELLSTATE_KEY, SHELLSTATE_VALUE,
            SHELLSTATE_CONFIRM_BYTE_INDEX, SHELLSTATE_CONFIRM_OFF,
        )
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SHELLSTATE_KEY) as key:
            value, vtype = winreg.QueryValueEx(key, SHELLSTATE_VALUE)
            if (vtype == winreg.REG_BINARY
                    and isinstance(value, (bytes, bytearray))
                    and len(value) >= SHELLSTATE_CONFIRM_BYTE_INDEX + 1):
                s['delete_confirm_enabled'] = (
                    value[SHELLSTATE_CONFIRM_BYTE_INDEX] != SHELLSTATE_CONFIRM_OFF
                )
            else:
                s['delete_confirm_enabled'] = True
    except (FileNotFoundError, OSError):
        s['delete_confirm_enabled'] = True
    except Exception as e:
        logger.log(f"collect delete_confirm error: {e}", module=MOD)
        s['delete_confirm_enabled'] = True

    try:
        from explorer import IMAGE_DETAILS_KEY, IMAGE_DETAILS_VALUE
        rr = reg_utils.reg('query', IMAGE_DETAILS_KEY, '/v', IMAGE_DETAILS_VALUE,
                           log=False)
        hidden = False
        if rr.returncode == 0 and IMAGE_DETAILS_VALUE in (rr.stdout or ''):
            m = re.search(r'REG_DWORD\s+(0x[0-9a-fA-F]+)', rr.stdout)
            if m and m.group(1).lower() == '0x1':
                hidden = True
        s['image_details_hidden'] = hidden
    except Exception as e:
        logger.log(f"collect image_details error: {e}", module=MOD)

    try:
        from explorer import FOLDERS, NAMESPACE_KEY
        for guid in FOLDERS:
            rr = reg_utils.reg('query', f'{NAMESPACE_KEY}\\{guid}', log=False)
            s[f'folder_{guid}'] = (rr.returncode == 0)
    except Exception as e:
        logger.log(f"collect folders error: {e}", module=MOD)
    return s


def collect_createnew():
    settings = {}

    if _key_exists(r'Folder\ShellNew', 'HKCR'):
        settings['folder'] = True
    elif _key_exists(r'Folder\ShellNew_no', 'HKCR'):
        settings['folder'] = False
    else:
        settings['folder'] = False

    extensions = {
        'bmp': r'SOFTWARE\Classes\.bmp\ShellNew',
        'contact': r'SOFTWARE\Classes\.contact\ShellNew',
        'lnk': r'SOFTWARE\Classes\.lnk\ShellNew',
        'rtf': r'SOFTWARE\Classes\.rtf\ShellNew',
        'txt': r'SOFTWARE\Classes\.txt\ShellNew',
        'zip': r'SOFTWARE\Classes\.zip\CompressedFolder\ShellNew',
    }
    for ext, path in extensions.items():
        if _key_exists(path, 'HKLM'):
            settings[ext] = True
        elif _key_exists(path + '_no', 'HKLM'):
            settings[ext] = False
        else:
            settings[ext] = False

    try:
        from contextmenu_pkg.createnew import (
            MS_EXTENSIONS, _ext_registered,
            _find_shellnew_under_ext, _classify_provider,
        )
        import winreg as _wr
        any_installed = False
        for ext in MS_EXTENSIONS:
            if not _ext_registered(ext):
                continue
            for hive in (_wr.HKEY_LOCAL_MACHINE, _wr.HKEY_CURRENT_USER):
                for h, path, cls, enabled in _find_shellnew_under_ext(hive, ext):
                    pid, _display, _vendor = _classify_provider(cls, ext)
                    any_installed = True
                    settings[f'provider_{pid}'] = enabled
        settings['office_installed'] = any_installed
    except Exception:
        settings['office_installed'] = False
    return settings


def collect_sendto():
    sendto_path = os.path.join(os.environ.get('APPDATA', ''),
                               'Microsoft', 'Windows', 'SendTo')
    settings = {}
    if not os.path.isdir(sendto_path):
        return settings
    try:
        import ctypes
        for name in sorted(os.listdir(sendto_path)):
            if name.lower() == 'desktop.ini':
                continue
            full = os.path.join(sendto_path, name)
            if not (os.path.isfile(full) or os.path.islink(full)):
                continue
            attrs = ctypes.windll.kernel32.GetFileAttributesW(full)
            if attrs == -1:
                continue
            settings[name] = not bool(attrs & 0x2)
    except Exception:
        pass
    return settings


# ---- save / load ----

def save_settings_to_ini(version, parent=None):
    try:
        path = filedialog.asksaveasfilename(
            title="Сохранить настройки",
            defaultextension=".ini",
            filetypes=[("INI files", "*.ini")],
            initialfile=f"interface_master_{datetime.now().strftime('%Y%m%d_%H%M%S')}.ini",
            parent=parent)
        if not path:
            return None
        config = ConfigParser()
        config.optionxform = str
        for section, values in collect_all_settings(version).items():
            config[section] = {}
            for key, value in values.items():
                config[section][key] = str(value).replace('%', '%%')
        with open(path, 'w', encoding='utf-8') as f:
            config.write(f)
        logger.log(f"saved settings to {path}", module=MOD)
        return path
    except Exception as e:
        logger.log(f"ERROR save_settings: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{e}")
        return None


def load_settings_from_ini(parent=None):
    """Возвращает (ConfigParser, path) или (None, None) если отменено/ошибка."""
    try:
        path = filedialog.askopenfilename(
            title="Загрузить настройки",
            filetypes=[("INI files", "*.ini")],
            parent=parent)
        if not path:
            return None, None
        config = ConfigParser()
        config.optionxform = str
        config.read(path, encoding='utf-8')
        logger.log(f"load settings from {path}", module=MOD)
        return config, path
    except Exception as e:
        logger.log(f"ERROR load_settings: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Не удалось загрузить настройки:\n{e}")
        return None, None