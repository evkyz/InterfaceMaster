import hashlib
import os
import re
import subprocess
import sys
import time
import tkinter as tk
import winreg
from configparser import ConfigParser
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

RECYCLE_PATH = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon"
DRIVE_ICONS = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"
CONTROL_PANEL_CLSIDS = [
    '{26EE0668-A00A-44D7-9371-BEB064C98683}',
    '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}',
    '{21EC2020-3AEA-1069-A2DD-08002B30309D}',
]
DESKTOP_ICON_PATHS = {
    'user': r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon",
    'computer': r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon",
    'network': r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon",
    'libraries': r"HKCR\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon",
}

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW

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

MOD_EXPECTED_HASHES = {
    "imageres.dll": "DDD5D7462A26F7837CD02A7667615354F9377A0A727BF880D825676D6DDC4CB7",
    "imagesp1.dll": "466E64D491C7B3D7051FBDC0907CD813D7A6435D18BF4B00CE0C3191B80E0541",
    "networkexplorer.dll": "7BF68211CFBDDE2E377C992CEC2F5A0C1F41B1F12BE3A99733A549AB63A8C961",
    "DDORes.dll": "4802C8F0BAAE3765FDFF16F09F56647B1C8B20889E4C531B773442EC068CBC3B",
    "mmres.dll": "4651BB26196D0ED493A26F4E401C8FAC0465E7F9DE2ECFED1642155A6EAF3703",
}

_HIVE_MAP = {
    'HKCU': winreg.HKEY_CURRENT_USER,
    'HKLM': winreg.HKEY_LOCAL_MACHINE,
    'HKCR': winreg.HKEY_CLASSES_ROOT,
}


def _reg(*args):
    return subprocess.run(['reg', *args], capture_output=True, text=True,
                          encoding='cp866', creationflags=CREATE_NO_WINDOW)


class AboutModule:
    IMASTER_MARKER = 'imaster.dll'

    def __init__(self, parent, version, dll_info=None):
        self.parent = parent
        self.version = version
        self.dll_info = dll_info or self._load_dll_info()
        self.frame = ttk.Frame(parent)
        self._create_widgets()
        self.show()

    @staticmethod
    def _load_dll_info():
        try:
            from main import DLLChecker
            return DLLChecker.get_dll_info()
        except ImportError:
            return {
                'expected_sha1': 'F62A4326A8300D9B828744E24314B98433510195',
                'actual_hash': None,
                'dll_path': None,
                'hash_match': False,
                'dll_found': False,
            }

    def _dll_status(self):
        if not self.dll_info['dll_found']:
            return "Не найден", '#e74c3c'
        if self.dll_info['hash_match']:
            return "Совпадает", '#27ae60'
        return "Не совпадает", '#e67e22'

    def _serialize_icon(self, value):
        if not value:
            return value
        s = str(value)
        low = s.lower()
        marker_pos = low.find(self.IMASTER_MARKER)
        if marker_pos == -1:
            return s
        return s[marker_pos:]

    def _deserialize_icon(self, value, current_dll_path):
        if not value:
            return value
        s = str(value)
        low = s.lower()
        if low.startswith(self.IMASTER_MARKER):
            if not current_dll_path:
                return s
            suffix = s[len(self.IMASTER_MARKER):]
            return f"{current_dll_path}{suffix}"
        marker_pos = low.find(self.IMASTER_MARKER)
        if marker_pos != -1 and current_dll_path:
            suffix = s[marker_pos + len(self.IMASTER_MARKER):]
            return f"{current_dll_path}{suffix}"
        return s

    def _is_imaster_icon(self, value):
        if not value:
            return False
        return self.IMASTER_MARKER in str(value).lower()

    @staticmethod
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

    @staticmethod
    def _is_admin():
        try:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    def _create_widgets(self):
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        ttk.Label(container, text="О программе",
                  font=('Arial', 18, 'bold'), foreground='#2c3e50').pack(pady=1)
        ttk.Label(container, text=f"Interface Master версия {self.version}",
                  font=('Arial', 10), foreground='#3498db').pack(pady=(0, 15))

        info = ttk.Frame(container)
        info.pack(fill=tk.BOTH, expand=True)

        self._section(info, "Описание",
                      "Программа для настройки интерфейса Windows\n"
                      "Помощник в кастомизации внешнего вида системы",
                      center=True)

        self._section(info, "Что можно настроить",
                      "• Иконки рабочего стола          • Иконки дисков\n"
                      "• Стрелки ярлыков                    • Контекстное меню\n"
                      "• Панель задач                           • Проводник\n"
                      "                            • Модификации",
                      center=False)

        self._dll_section(info)

        self._section(info, "Разработка",
                      "Идея: EvKyz\nРеализация: DeepSeek AI\n"
                      "Сборка: 10 ноября 2026\nЛицензия: MIT License",
                      font=('Arial', 8), fg='#7f8c8d', center=True)

        self._buttons(container)

    def _section(self, parent, title, text, font=('Arial', 9), fg=None, center=True):
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.pack(fill=tk.X, pady=(0, 10))
        lbl = ttk.Label(frame, text=text, font=font,
                        justify=tk.CENTER if center else tk.LEFT)
        if fg:
            lbl.config(foreground=fg)
        lbl.pack()

    def _dll_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Файл imaster.dll", padding=10)
        frame.pack(fill=tk.X, pady=(0, 15))
        status_text, color = self._dll_status()
        tk.Label(frame, text=f"Статус: {status_text}",
                 font=('Consolas', 9, 'bold'), fg=color, anchor='w').pack(
            fill=tk.X, pady=(0, 5))
        if self.dll_info['dll_found'] and self.dll_info['dll_path']:
            tk.Label(frame, text=f"Путь: {self.dll_info['dll_path']}",
                     font=('Consolas', 8), anchor='w').pack(fill=tk.X, pady=(0, 3))
            if self.dll_info['actual_hash']:
                tk.Label(frame, text=f"Хеш SHA1: {self.dll_info['actual_hash']}",
                         font=('Consolas', 8), anchor='w').pack(fill=tk.X)
        else:
            tk.Label(frame, text="Файл imaster.dll не найден",
                     font=('Consolas', 8), anchor='w').pack(fill=tk.X)

    def _buttons(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, pady=(10, 5))
        left = ttk.Frame(bar)
        left.pack(side=tk.LEFT)

        save_btn = ttk.Button(left, text="Сохранить настройки",
                              command=self.save_settings_to_ini, width=20)
        save_btn.pack(side=tk.LEFT, padx=(0, 5))
        self._tooltip(save_btn, "Сохранить все текущие настройки в INI файл")

        load_btn = ttk.Button(left, text="Загрузить настройки",
                              command=self.load_settings_from_ini, width=20)
        load_btn.pack(side=tk.LEFT, padx=(0, 5))
        self._tooltip(load_btn, "Загрузить настройки из INI файла и применить их")

        reset_btn = ttk.Button(bar, text="Сбросить все настройки",
                               command=self.reset_all_settings, width=25)
        reset_btn.pack(side=tk.RIGHT)
        self._tooltip(reset_btn,
                      "Сбросить все настройки программы к стандартным значениям Windows")

    def _tooltip(self, widget, text):
        state = {'tip': None}

        def enter(event):
            if state['tip']:
                return
            try:
                tip = tk.Toplevel()
                tip.wm_overrideredirect(True)
                tip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
                tk.Label(tip, text=text, background="lightyellow", relief="solid",
                         borderwidth=1, font=("Arial", 8)).pack()
                state['tip'] = tip
            except Exception:
                state['tip'] = None

        def leave(event):
            if state['tip']:
                try:
                    state['tip'].destroy()
                except Exception:
                    pass
                state['tip'] = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    @staticmethod
    def _get_registry_value(key_path, value_name='', hive='HKCU'):
        try:
            with winreg.OpenKey(_HIVE_MAP[hive], key_path) as k:
                value, vtype = winreg.QueryValueEx(k, value_name)
                if vtype == winreg.REG_DWORD:
                    return hex(value)
                return value
        except (FileNotFoundError, OSError, KeyError):
            return None

    @staticmethod
    def _key_exists(key_path, hive='HKLM'):
        try:
            with winreg.OpenKey(_HIVE_MAP[hive], key_path):
                return True
        except (FileNotFoundError, OSError, KeyError):
            return False

    def collect_all_settings(self):
        return {
            'metadata': {
                'version': self.version,
                'export_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'program': 'Interface Master',
            },
            'desktop': self._collect_desktop(),
            'disk': self._collect_disk(),
            'str': self._collect_shortcut(),
            'contextmenu': self._collect_contextmenu(),
            'taskbar': self._collect_taskbar(),
            'explorer': self._collect_explorer(),
            'createnew': self._collect_createnew(),
            'sendto': self._collect_sendto(),
        }

    def _collect_desktop(self):
        settings = {}
        items = {
            'user': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon',
            'computer': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon',
            'network': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon',
        }
        for item, path in items.items():
            val = self._get_registry_value(path, '', 'HKCU')
            if val and self._is_imaster_icon(val):
                settings[f'icon_{item}'] = self._serialize_icon(val)

        recycle_reg = r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon'
        empty_default = self._get_registry_value(recycle_reg, '', 'HKCU')
        empty_explicit = self._get_registry_value(recycle_reg, 'Empty', 'HKCU')
        full_val = self._get_registry_value(recycle_reg, 'Full', 'HKCU')

        src_empty = empty_explicit if (empty_explicit and self._is_imaster_icon(empty_explicit)) else empty_default
        if src_empty and self._is_imaster_icon(src_empty):
            matched = False
            if ',-' in src_empty:
                idx = src_empty.split(',-')[-1]
                if idx in RECYCLE_INDEX_PAIRS:
                    empty_idx, full_idx = RECYCLE_INDEX_PAIRS[idx]
                    dll = src_empty.split(',-')[0]
                    settings['icon_recycle_bin_empty'] = self._serialize_icon(
                        f"{dll},-{empty_idx}")
                    if full_val and self._is_imaster_icon(full_val):
                        settings['icon_recycle_bin_full'] = self._serialize_icon(full_val)
                    matched = True
            if not matched:
                settings['icon_recycle_bin_empty'] = self._serialize_icon(src_empty)
                if full_val and self._is_imaster_icon(full_val):
                    settings['icon_recycle_bin_full'] = self._serialize_icon(full_val)

        libraries = self._get_registry_value(
            r'CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon', '', 'HKCR')
        if libraries and self._is_imaster_icon(libraries):
            settings['icon_libraries'] = self._serialize_icon(libraries)

        for clsid in CONTROL_PANEL_CLSIDS:
            path = rf'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{clsid}\DefaultIcon'
            val = self._get_registry_value(path, '', 'HKCU')
            if val and self._is_imaster_icon(val):
                settings[f'icon_control_panel_{clsid}'] = self._serialize_icon(val)
        return settings

    def _collect_disk(self):
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
                    if val and self._is_imaster_icon(val):
                        settings[f'icon_{drive.upper()}'] = self._serialize_icon(val)
        except FileNotFoundError:
            pass
        return settings

    def _collect_shortcut(self):
        settings = {}
        arrow = self._get_registry_value(
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons', '29', 'HKLM')
        if arrow is None:
            settings['arrow_style'] = 'default'
        elif arrow == '':
            settings['arrow_style'] = 'none'
        elif 'shell32.dll,-30' in str(arrow):
            settings['arrow_style'] = 'small'
        else:
            settings['arrow_style'] = 'custom'

        tmpl = self._get_registry_value(
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

    def _collect_contextmenu(self):
        settings = {}
        try:
            from contextmenu import check_menu_item_exists, CONTEXT_MENU_KEYS
            for key, item_type in CONTEXT_MENU_KEYS.items():
                try:
                    settings[key] = bool(check_menu_item_exists(item_type))
                except Exception:
                    settings[key] = False
        except Exception as e:
            self._log(f"  ОШИБКА сбора contextmenu: {e}")
        return settings

    def _collect_taskbar(self):
        s = {}
        short_date = self._get_registry_value(r'Control Panel\International',
                                              'sShortDate', 'HKCU')
        s['show_weekday'] = 'ddd' in str(short_date) if short_date else False

        notif = self._get_registry_value(
            r'SOFTWARE\Policies\Microsoft\Windows\Explorer',
            'DisableNotificationCenter', 'HKLM')
        s['notification_center'] = notif != '0x1'

        small = self._get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'TaskbarSmallIcons', 'HKCU')
        s['small_icons'] = small == '0x1'

        search = self._get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Search',
            'SearchboxTaskbarMode', 'HKCU')
        s['search_mode'] = {'0x0': 0, '0x1': 1}.get(search, 2)

        tv = self._get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'ShowTaskViewButton', 'HKCU')
        s['show_task_view'] = tv != '0x0'

        glom = self._get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'TaskbarGlomLevel', 'HKCU')
        s['grouping_mode'] = {'0x0': 0, '0x1': 1, '0x2': 2}.get(glom, 0)

        try:
            from taskbar import is_people_bar_supported, is_windows_version_supported
            if is_people_bar_supported():
                people = self._get_registry_value(
                    r'Software\Policies\Microsoft\Windows\Explorer',
                    'HidePeopleBar', 'HKCU')
                s['show_people'] = people != '0x1'
            if is_windows_version_supported():
                meet = self._get_registry_value(
                    r'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer',
                    'HideSCAMeetNow', 'HKLM')
                s['show_meet_now'] = meet != '0x1'
        except Exception as e:
            self._log(f"  ОШИБКА сбора taskbar (people/meet): {e}")
        return s

    def _collect_explorer(self):
        s = {}
        # BitLocker: ProgrammaticAccessOnly — значение в HKCR\Drive\shell\encrypt-bde.
        # reg.exe не понимает HKCR — используем winreg напрямую.
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                                r'Drive\shell\encrypt-bde') as k:
                winreg.QueryValueEx(k, 'ProgrammaticAccessOnly')
                s['bitlocker_enabled'] = False
        except (FileNotFoundError, OSError):
            s['bitlocker_enabled'] = True

        photo_assoc = self._get_registry_value(
            r'SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
            '.jpg', 'HKLM')
        photo_class = self._key_exists(
            r'SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg', 'HKLM')
        s['photo_viewer_enabled'] = (photo_assoc is not None) and photo_class

        hide_ext = self._get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'HideFileExt', 'HKCU')
        s['show_file_extensions'] = hide_ext == '0x0'

        no_recent = self._get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer',
            'NoRecentDocsHistory', 'HKCU')
        s['recent_folders_enabled'] = no_recent != '0x1'

        try:
            from explorer import IMAGE_DETAILS_KEY, IMAGE_DETAILS_VALUE
            rr = _reg('query', IMAGE_DETAILS_KEY, '/v', IMAGE_DETAILS_VALUE)
            hidden = False
            if rr.returncode == 0 and IMAGE_DETAILS_VALUE in rr.stdout:
                m = re.search(r'REG_DWORD\s+(0x[0-9a-fA-F]+)', rr.stdout)
                if m and m.group(1).lower() == '0x1':
                    hidden = True
            s['image_details_hidden'] = hidden
        except Exception as e:
            self._log(f"  ОШИБКА сбора image_details: {e}")

        try:
            from explorer import FOLDERS, NAMESPACE_KEY
            for guid in FOLDERS:
                rr = _reg('query', f'{NAMESPACE_KEY}\\{guid}')
                s[f'folder_{guid}'] = (rr.returncode == 0)
        except Exception as e:
            self._log(f"  ОШИБКА сбора папок explorer: {e}")
        return s

    def _collect_createnew(self):
        settings = {}

        if self._key_exists(r'Folder\ShellNew', 'HKCR'):
            settings['folder'] = True
        elif self._key_exists(r'Folder\ShellNew_no', 'HKCR'):
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
            if self._key_exists(path, 'HKLM'):
                settings[ext] = True
            elif self._key_exists(path + '_no', 'HKLM'):
                settings[ext] = False
            else:
                settings[ext] = False

        try:
            from createnew import (
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

    def _collect_sendto(self):
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
                hidden = bool(attrs & 0x2)
                settings[name] = not hidden
        except Exception:
            pass
        return settings

    def save_settings_to_ini(self):
        try:
            path = filedialog.asksaveasfilename(
                title="Сохранить настройки",
                defaultextension=".ini",
                filetypes=[("INI files", "*.ini")],
                initialfile=f"interface_master_{datetime.now().strftime('%Y%m%d_%H%M%S')}.ini")
            if not path:
                return
            config = ConfigParser()
            config.optionxform = str
            for section, values in self.collect_all_settings().items():
                config[section] = {}
                for key, value in values.items():
                    config[section][key] = str(value).replace('%', '%%')
            with open(path, 'w', encoding='utf-8') as f:
                config.write(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{e}")

    def load_settings_from_ini(self):
        try:
            path = filedialog.askopenfilename(title="Загрузить настройки",
                                              filetypes=[("INI files", "*.ini")])
            if not path:
                return
            config = ConfigParser()
            config.optionxform = str
            config.read(path, encoding='utf-8')

            self._log("=" * 60)
            self._log("ЗАГРУЗКА НАСТРОЕК ИЗ INI")
            for sec in ('str', 'createnew', 'sendto', 'taskbar',
                        'contextmenu', 'explorer'):
                if sec in config:
                    self._log(f"Секция [{sec}]:")
                    for k, v in config[sec].items():
                        self._log(f"  {k} = {v}")
            self._log("-" * 60)

            self._apply_settings(config)
            self._restart_explorer_silent()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить настройки:\n{e}")

    def _ini_has_imaster_icons(self, config):
        for sec in ('desktop', 'disk'):
            if sec not in config:
                continue
            for key, value in config[sec].items():
                if key.startswith('icon_') and self._is_imaster_icon(value):
                    return True
        return False

    def _ensure_dll_path_for_ini(self, config):
        """Определяет DLL-путь для загрузки desktop/disk из INI.

        Возвращает (dll_path, has_imaster):
        - (None, False) — imaster-иконок в INI нет → применяем desktop/disk как есть.
        - (None, True) — imaster-иконки есть, но DLL не найдена нигде → пропускаем
          desktop/disk.
        - (path, True) — путь найден/скопирован → применяем desktop/disk.

        Логика:
        - Если путь уже сохранён в реестре и валиден — используем его без диалога.
        - Если в этой сессии уже предлагали копирование — используем recommended_path.
        - Иначе делаем свежую проверку DLL в папке exe / System32 / др.
          (не полагаемся на dll_info — он мог устареть, если DLL появилась после старта).
        - Если DLL физически есть — предлагаем скопировать в System32 (диалог).
        - Если DLL физически нет — пропускаем desktop/disk.
        """
        has_imaster = self._ini_has_imaster_icons(config)
        if not has_imaster:
            return None, False

        try:
            from main import DLLChecker, DLLCheckerInternal
            from main import calculate_sha1, VALID_HASHES
        except ImportError:
            return None, True

        # 1) Путь уже сохранён в реестре и валиден — используем без диалога.
        saved = DLLChecker.get_dll_path_from_registry()
        if saved and os.path.isfile(saved):
            h = calculate_sha1(saved)
            if h and h in VALID_HASHES:
                return saved, True

        # 2) В этой сессии уже предлагали копирование — используем recommended_path.
        if DLLChecker.copy_offered:
            status = DLLChecker.get_dll_status()
            return status.get('recommended_path'), True

        # 3) Свежая проверка: есть ли DLL в папке exe / System32 / других местах.
        base = sys._MEIPASS if getattr(sys, 'frozen', False) \
            else os.path.dirname(os.path.abspath(__file__))
        fresh = DLLCheckerInternal(base).check_imaster_dll()
        if not fresh.get('dll_found') or not fresh.get('hash_match'):
            # DLL физически нет — пропускаем desktop/disk
            return None, True

        # DLL есть, но путь не сохранён → предлагаем копирование в System32
        path = DLLChecker.offer_copy_to_system32(self.parent, "desktop")
        return path, True

    def _apply_settings(self, config):
        if 'str' in config:
            self._apply_str(config)
        if 'taskbar' in config:
            self._apply_taskbar(config)
        if 'explorer' in config:
            self._apply_explorer(config)

        dll_path, has_imaster = self._ensure_dll_path_for_ini(config)
        skip_imaster_sections = has_imaster and dll_path is None

        if 'desktop' in config:
            if skip_imaster_sections:
                self._log("Секция [desktop]: пропущена (imaster.dll не найдена)")
            else:
                self._apply_desktop(config, dll_path=dll_path)

        if 'disk' in config:
            if skip_imaster_sections:
                self._log("Секция [disk]: пропущена (imaster.dll не найдена)")
            else:
                self._apply_disk(config, dll_path=dll_path)

        if 'contextmenu' in config:
            self._apply_contextmenu(config)
        if 'createnew' in config:
            self._apply_createnew(config)
        if 'sendto' in config:
            self._apply_sendto(config)
        self._log("=" * 60)
        self._log("ПРИМЕНЕНИЕ НАСТРОЕК ЗАВЕРШЕНО")
        self._log("=" * 60)

    def _apply_str(self, config):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК ЯРЛЫКОВ ===")
        try:
            from str import Module as StrModule
            temp = ttk.Frame(self.parent)
            s = StrModule(temp)
            temp.pack_forget()
            changed = False
            if 'arrow_style' in config['str']:
                style = config['str']['arrow_style']
                self._log(f"  arrow_style = {style}")
                mapping = {'default': 1, 'small': 2, 'none': 3}
                if style in mapping:
                    s.arrow_choice.set(mapping[style])
                    changed = True
            if 'shortcut_text' in config['str']:
                txt = config['str']['shortcut_text']
                self._log(f"  shortcut_text = {txt}")
                if txt in ('default', 'removed'):
                    s.shortcut_text_choice.set(1 if txt == 'default' else 2)
                    changed = True
            if changed:
                s.apply_changes()
                self._log("    Изменения применены")
            s.hide()
            temp.destroy()
        except Exception as e:
            self._log(f"  ОШИБКА str: {e}")

    def _apply_taskbar(self, config):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК ПАНЕЛИ ЗАДАЧ ===")
        try:
            from taskbar import Module as TaskbarModule
            temp = ttk.Frame(self.parent)
            t = TaskbarModule(temp)
            tb = config['taskbar']

            def _target_bool(key):
                return tb[key].strip().lower() == 'true' if key in tb else None

            def _target_int(key):
                try:
                    return int(tb[key].strip()) if key in tb else None
                except (ValueError, TypeError):
                    return None

            applied = 0
            skipped = 0
            failed = 0

            tgt = _target_bool('show_weekday')
            if tgt is not None:
                cur = t.check_weekday_state()
                if cur != tgt:
                    t._do_toggle_weekday()
                    if t.check_weekday_state() == tgt:
                        self._log(f"  show_weekday: {cur} -> {tgt} (OK)")
                        applied += 1
                    else:
                        self._log(f"  show_weekday: не удалось применить")
                        failed += 1
                else:
                    self._log(f"  show_weekday: без изменений ({tgt})")

            tgt = _target_bool('small_icons')
            if tgt is not None:
                cur = t.check_taskbar_size_state()
                if cur != tgt:
                    t._do_toggle_taskbar_size()
                    if t.check_taskbar_size_state() == tgt:
                        self._log(f"  small_icons: {cur} -> {tgt} (OK)")
                        applied += 1
                    else:
                        self._log(f"  small_icons: не удалось применить")
                        failed += 1
                else:
                    self._log(f"  small_icons: без изменений ({tgt})")

            tgt = _target_int('search_mode')
            if tgt is not None and tgt in (0, 1, 2):
                for _ in range(5):
                    if t.check_search_state() == tgt:
                        break
                    t._do_toggle_search()
                final = t.check_search_state()
                if final == tgt:
                    self._log(f"  search_mode: -> {final} (OK)")
                    applied += 1
                else:
                    self._log(f"  search_mode: не удалось ({final} != {tgt})")
                    failed += 1

            tgt = _target_bool('show_task_view')
            if tgt is not None:
                cur = not t.check_task_view_state()
                if cur != tgt:
                    t._do_toggle_task_view()
                    if (not t.check_task_view_state()) == tgt:
                        self._log(f"  show_task_view: {cur} -> {tgt} (OK)")
                        applied += 1
                    else:
                        self._log(f"  show_task_view: не удалось")
                        failed += 1
                else:
                    self._log(f"  show_task_view: без изменений ({tgt})")

            tgt = _target_int('grouping_mode')
            if tgt is not None and tgt in (0, 1, 2):
                for _ in range(5):
                    if t.check_grouping_state() == tgt:
                        break
                    t._do_toggle_grouping()
                final = t.check_grouping_state()
                if final == tgt:
                    self._log(f"  grouping_mode: -> {final} (OK)")
                    applied += 1
                else:
                    self._log(f"  grouping_mode: не удалось ({final} != {tgt})")
                    failed += 1

            is_admin = self._is_admin()

            tgt = _target_bool('notification_center')
            if tgt is not None:
                if not is_admin:
                    self._log(f"  notification_center: пропущено (нужны права администратора)")
                    skipped += 1
                else:
                    cur = not t.check_notification_center_state()
                    if cur != tgt:
                        t._do_toggle_notification_center()
                        if (not t.check_notification_center_state()) == tgt:
                            self._log(f"  notification_center: {cur} -> {tgt} (OK)")
                            applied += 1
                        else:
                            self._log(f"  notification_center: не удалось применить")
                            failed += 1
                    else:
                        self._log(f"  notification_center: без изменений ({tgt})")

            tgt = _target_bool('show_people')
            if tgt is not None:
                if not t.is_people_bar_supported():
                    self._log(f"  show_people: пропущено (не поддерживается на этой версии)")
                    skipped += 1
                else:
                    cur = not t.check_people_state()
                    if cur != tgt:
                        t._do_toggle_people()
                        if (not t.check_people_state()) == tgt:
                            self._log(f"  show_people: {cur} -> {tgt} (OK)")
                            applied += 1
                        else:
                            self._log(f"  show_people: не удалось применить")
                            failed += 1
                    else:
                        self._log(f"  show_people: без изменений ({tgt})")

            tgt = _target_bool('show_meet_now')
            if tgt is not None:
                if not t.is_windows_version_supported():
                    self._log(f"  show_meet_now: пропущено (не поддерживается на этой версии)")
                    skipped += 1
                elif not is_admin:
                    self._log(f"  show_meet_now: пропущено (нужны права администратора)")
                    skipped += 1
                else:
                    cur = not t.check_meet_now_state()
                    if cur != tgt:
                        t._do_toggle_meet_now()
                        if (not t.check_meet_now_state()) == tgt:
                            self._log(f"  show_meet_now: {cur} -> {tgt} (OK)")
                            applied += 1
                        else:
                            self._log(f"  show_meet_now: не удалось применить")
                            failed += 1
                    else:
                        self._log(f"  show_meet_now: без изменений ({tgt})")

            t.hide()
            temp.destroy()
            self._log(f"  Итог taskbar: применено {applied}, пропущено {skipped}, ошибок {failed}")
        except Exception as e:
            self._log(f"  ОШИБКА taskbar: {e}")

    def _apply_explorer(self, config):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК ПРОВОДНИКА ===")
        try:
            from explorer import ExplorerModule, FOLDERS, NAMESPACE_KEY
            temp = ttk.Frame(self.parent)
            e = ExplorerModule(temp)
            temp.pack_forget()
            ex = config['explorer']
            is_admin = self._is_admin()

            def _target(k):
                return ex[k].lower() == 'true' if k in ex else None

            applied = 0
            skipped = 0
            failed = 0

            if 'bitlocker_enabled' in ex:
                tgt = _target('bitlocker_enabled')
                if tgt != e.bitlocker_enabled:
                    if not is_admin:
                        self._log("  bitlocker_enabled: пропущено (нужны права администратора)")
                        skipped += 1
                    else:
                        e.toggle_bitlocker()
                        if e.bitlocker_enabled == tgt:
                            self._log(f"  bitlocker_enabled: -> {tgt} (OK)")
                            applied += 1
                        else:
                            self._log("  bitlocker_enabled: не удалось применить")
                            failed += 1
                else:
                    self._log(f"  bitlocker_enabled: без изменений ({tgt})")

            if 'photo_viewer_enabled' in ex:
                tgt = _target('photo_viewer_enabled')
                if tgt != e.photo_viewer_enabled:
                    if not is_admin:
                        self._log("  photo_viewer_enabled: пропущено (нужны права администратора)")
                        skipped += 1
                    else:
                        e.toggle_photo_viewer()
                        if e.photo_viewer_enabled == tgt:
                            self._log(f"  photo_viewer_enabled: -> {tgt} (OK)")
                            applied += 1
                        else:
                            self._log("  photo_viewer_enabled: не удалось применить")
                            failed += 1
                else:
                    self._log(f"  photo_viewer_enabled: без изменений ({tgt})")

            if 'show_file_extensions' in ex:
                tgt = _target('show_file_extensions')
                if tgt != e.file_ext_enabled:
                    e.toggle_file_ext()
                    if e.file_ext_enabled == tgt:
                        self._log(f"  show_file_extensions: -> {tgt} (OK)")
                        applied += 1
                    else:
                        self._log("  show_file_extensions: не удалось применить")
                        failed += 1
                else:
                    self._log(f"  show_file_extensions: без изменений ({tgt})")

            if 'recent_folders_enabled' in ex:
                tgt = _target('recent_folders_enabled')
                if tgt != e.recent_folders_enabled:
                    e.toggle_recent_folders()
                    if e.recent_folders_enabled == tgt:
                        self._log(f"  recent_folders_enabled: -> {tgt} (OK)")
                        applied += 1
                    else:
                        self._log("  recent_folders_enabled: не удалось применить")
                        failed += 1
                else:
                    self._log(f"  recent_folders_enabled: без изменений ({tgt})")

            if 'image_details_hidden' in ex:
                tgt = _target('image_details_hidden')
                supported = getattr(e, 'image_details_supported', False)
                cur = getattr(e, 'image_details_hidden', None)
                if not supported:
                    self._log("  image_details_hidden: пропущено (не поддерживается на этой версии)")
                    skipped += 1
                elif cur is None:
                    self._log("  image_details_hidden: состояние недоступно")
                    skipped += 1
                elif tgt != cur:
                    e.toggle_image_details()
                    if e.image_details_hidden == tgt:
                        self._log(f"  image_details_hidden: -> {tgt} (OK)")
                        applied += 1
                    else:
                        self._log("  image_details_hidden: не удалось применить")
                        failed += 1
                else:
                    self._log(f"  image_details_hidden: без изменений ({tgt})")

            for guid in FOLDERS:
                key = f'folder_{guid}'
                if key not in ex:
                    continue
                tgt = _target(key)
                reg_path = f"{NAMESPACE_KEY}\\{guid}"
                cur = _reg('query', reg_path).returncode == 0
                if tgt == cur:
                    self._log(f"  {key}: без изменений ({tgt})")
                    skipped += 1
                    continue
                if not is_admin:
                    self._log(f"  {key}: пропущено (нужны права администратора)")
                    skipped += 1
                    continue
                if tgt:
                    _reg('add', reg_path, '/f')
                    ok = _reg('query', reg_path).returncode == 0
                else:
                    _reg('delete', reg_path, '/f')
                    ok = _reg('query', reg_path).returncode != 0
                if ok:
                    self._log(f"  {key}: -> {tgt} (OK)")
                    applied += 1
                else:
                    self._log(f"  {key}: не удалось применить")
                    failed += 1

            e.hide()
            temp.destroy()
            self._log(f"  Итог explorer: применено {applied}, "
                      f"без изменений {skipped}, ошибок {failed}")
        except Exception as e:
            self._log(f"  ОШИБКА explorer: {e}")

    def _apply_desktop(self, config, dll_path=None):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК ИКОНОК РАБОЧЕГО СТОЛА ===")
        if dll_path is None:
            dll_path = self._get_current_imaster_path()
        if dll_path:
            self._log(f"  Текущий imaster.dll: {dll_path}")
        else:
            self._log("  imaster.dll не найден — пути imaster.dll останутся "
                      "как в INI (могут не работать)")

        applied = 0
        skipped_system = 0
        for key, value in config['desktop'].items():
            if not key.startswith('icon_'):
                continue
            icon_type = key.replace('icon_', '')

            if not self._is_imaster_icon(value):
                self._log(f"  {icon_type}: пропущено (системная иконка, "
                          f"не imaster.dll): {value}")
                skipped_system += 1
                continue

            real_value = self._deserialize_icon(value, dll_path)
            self._log(f"  {icon_type} = {value} -> {real_value}")

            if icon_type in ('user', 'computer', 'network'):
                _reg('add', DESKTOP_ICON_PATHS[icon_type], '/ve', '/d', real_value, '/f')
            elif icon_type == 'recycle_bin_empty':
                _reg('add', RECYCLE_PATH, '/ve', '/d', real_value, '/f')
                _reg('add', RECYCLE_PATH, '/v', 'Empty', '/d', real_value, '/f')
            elif icon_type == 'recycle_bin_full':
                _reg('add', RECYCLE_PATH, '/v', 'Full', '/d', real_value, '/f')
            elif icon_type == 'libraries':
                ok = False
                for hive, path in (
                    (winreg.HKEY_LOCAL_MACHINE,
                     r"SOFTWARE\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
                    (winreg.HKEY_CURRENT_USER,
                     r"Software\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
                    (winreg.HKEY_CLASSES_ROOT,
                     r"CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
                ):
                    try:
                        try:
                            k = winreg.OpenKey(hive, path, 0, winreg.KEY_WRITE)
                        except FileNotFoundError:
                            k = winreg.CreateKey(hive, path)
                        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, real_value)
                        winreg.CloseKey(k)
                        ok = True
                        break
                    except (PermissionError, OSError):
                        continue
                    except Exception:
                        continue
                if not ok:
                    self._log(f"  ОШИБКА libraries: не удалось записать")
            elif icon_type == 'control_panel':
                for clsid in CONTROL_PANEL_CLSIDS:
                    path = (f"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\"
                            f"{clsid}\\DefaultIcon")
                    _reg('add', path, '/ve', '/d', real_value, '/f')
            elif icon_type.startswith('control_panel_'):
                clsid = icon_type.replace('control_panel_', '')
                path = (f"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\"
                        f"{clsid}\\DefaultIcon")
                _reg('add', path, '/ve', '/d', real_value, '/f')
            applied += 1

        if applied:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True,
                           creationflags=CREATE_NO_WINDOW)
        self._log(f"  Итог: применено {applied}, "
                  f"пропущено (системные) {skipped_system}")

    def _apply_disk(self, config, dll_path=None):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК ИКОНОК ДИСКОВ ===")
        if dll_path is None:
            dll_path = self._get_current_imaster_path()
        if dll_path:
            self._log(f"  Текущий imaster.dll: {dll_path}")
        else:
            self._log("  imaster.dll не найден — пути imaster.dll останутся "
                      "как в INI (могут не работать)")

        try:
            import ctypes
            bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        except Exception:
            bitmask = 0

        def _drive_exists(letter):
            idx = ord(letter.upper()) - ord('A')
            if 0 <= idx < 26:
                return bool(bitmask & (1 << idx))
            return False

        _reg('delete', DRIVE_ICONS, '/f')

        applied = 0
        skipped_missing = 0
        skipped_system = 0
        for key, value in config['disk'].items():
            if not (key.startswith('icon_') and value):
                continue
            drive = key.replace('icon_', '').upper()
            if not (len(drive) == 1 and drive.isalpha()):
                self._log(f"  {key}: некорректная буква диска, пропущено")
                skipped_missing += 1
                continue
            if not self._is_imaster_icon(value):
                self._log(f"  Диск {drive}: пропущено (системная иконка, "
                          f"не imaster.dll): {value}")
                skipped_system += 1
                continue
            if not _drive_exists(drive):
                self._log(f"  Диск {drive}: пропущен (диск отсутствует в системе)")
                skipped_missing += 1
                continue
            real_value = self._deserialize_icon(value, dll_path)
            path = f"{DRIVE_ICONS}\\{drive}\\DefaultIcon"
            _reg('add', path, '/ve', '/d', real_value, '/f')
            self._log(f"  Диск {drive}: {value} -> {real_value}")
            applied += 1

        self._log(f"  Итог: применено {applied}, "
                  f"пропущено (нет диска) {skipped_missing}, "
                  f"пропущено (системные) {skipped_system}")

    def _apply_contextmenu(self, config):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК КОНТЕКСТНОГО МЕНЮ ===")
        try:
            from contextmenu import DesktopMenuTab, ITEM_NAME_BY_KEY
            temp = ttk.Frame(self.parent)
            dt = DesktopMenuTab(temp)
            applied = 0
            skipped = 0
            not_found = 0
            for key, state in config['contextmenu'].items():
                target = state.lower() == 'true'
                menu_name = ITEM_NAME_BY_KEY.get(key, key)
                found = False
                for btn_info in dt.menu_item_buttons:
                    item = btn_info['data']
                    if item['name'] == menu_name:
                        found = True
                        cur = item['check_command']()
                        if cur != target:
                            (item['add_command'] if target else item['remove_command'])()
                            self._log(f"  {key}: {cur} -> {target}")
                            applied += 1
                        else:
                            self._log(f"  {key}: без изменений ({target})")
                            skipped += 1
                        break
                if not found:
                    self._log(f"  {key}: пункт '{menu_name}' не найден в UI")
                    not_found += 1
            temp.destroy()
            self._log(f"  Итог contextmenu: применено {applied}, "
                      f"без изменений {skipped}, не найдено {not_found}")
        except Exception as e:
            self._log(f"  ОШИБКА contextmenu: {e}")

    def _apply_createnew(self, config):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК МЕНЮ 'СОЗДАТЬ' ===")
        try:
            from createnew import (
                _reg_exists, _rename_key, _ensure_shellnew,
                TARGET_EXTENSIONS, MS_EXTENSIONS,
                _ext_registered, _find_shellnew_under_ext, _classify_provider,
            )
            import winreg as _wr
            cn = config['createnew']

            def _target(key):
                return cn[key].strip().lower() == 'true' if key in cn else None

            def _set_state(hive, base_path, enabled, label):
                try:
                    if enabled:
                        if _reg_exists(hive, base_path + '_no'):
                            _rename_key(hive, base_path + '_no', base_path)
                        elif not _reg_exists(hive, base_path):
                            _ensure_shellnew(hive, base_path)
                        self._log(f"  {label}: включено")
                    else:
                        if _reg_exists(hive, base_path):
                            _rename_key(hive, base_path, base_path + '_no')
                        self._log(f"  {label}: отключено")
                except Exception as e:
                    self._log(f"  ОШИБКА {label}: {e}")

            folder_state = _target('folder')
            if folder_state is not None:
                _set_state(_wr.HKEY_CLASSES_ROOT, r'Folder\ShellNew',
                           folder_state, 'folder')

            for ext, path in TARGET_EXTENSIONS.items():
                state = _target(ext.lstrip('.'))
                if state is None:
                    continue
                _set_state(_wr.HKEY_LOCAL_MACHINE, path, state, ext)

            ini_providers = {}
            for key, value in cn.items():
                if key.startswith('provider_'):
                    pid = key[len('provider_'):]
                    ini_providers[pid] = value.strip().lower() == 'true'

            installed = {}
            for ext in MS_EXTENSIONS:
                if not _ext_registered(ext):
                    continue
                for hive in (_wr.HKEY_LOCAL_MACHINE, _wr.HKEY_CURRENT_USER):
                    for h, path, cls, current_enabled in _find_shellnew_under_ext(hive, ext):
                        pid, _display, _vendor = _classify_provider(cls, ext)
                        if pid in installed:
                            continue
                        installed[pid] = (h, path, current_enabled, cls)

            if not ini_providers:
                self._log("  Офисные пункты: в INI нет ни одного provider_*, "
                          "пропущено всё")
            elif not installed:
                self._log("  Офисные пункты: в системе нет ни одного офисного "
                          "клиента, пропущено всё")
                self._log(f"    (в INI было: {', '.join(sorted(ini_providers))})")
            else:
                applied = 0
                skipped_not_installed = 0
                untouched = sorted(set(installed) - set(ini_providers))
                for pid, target in ini_providers.items():
                    if pid not in installed:
                        self._log(f"  {pid}: пропущено (не установлен)")
                        skipped_not_installed += 1
                        continue
                    hive, path, current_enabled, cls = installed[pid]
                    if target == current_enabled:
                        self._log(f"  {pid} ({cls}): без изменений "
                                  f"({'включено' if target else 'отключено'})")
                        continue
                    try:
                        if target:
                            if path.endswith('_no'):
                                _rename_key(hive, path, path[:-3])
                            self._log(f"  {pid} ({cls}): включено")
                        else:
                            if not path.endswith('_no'):
                                _rename_key(hive, path, path + '_no')
                            self._log(f"  {pid} ({cls}): отключено")
                        applied += 1
                    except Exception as e:
                        self._log(f"  ОШИБКА {pid}: {e}")

                if untouched:
                    self._log(f"  Не тронуты (есть в системе, но нет в INI): "
                              f"{', '.join(untouched)}")
                self._log(f"  Итог: применено {applied}, "
                          f"пропущено (нет в системе) {skipped_not_installed}")
        except Exception as e:
            self._log(f"  ОШИБКА createnew: {e}")

    def _apply_sendto(self, config):
        self._log("=== ПРИМЕНЕНИЕ НАСТРОЕК МЕНЮ 'ОТПРАВИТЬ' ===")
        try:
            import ctypes
            sendto_path = os.path.join(os.environ.get('APPDATA', ''),
                                       'Microsoft', 'Windows', 'SendTo')
            if not os.path.isdir(sendto_path):
                self._log("  Папка SendTo не найдена")
                return
            for name, value in config['sendto'].items():
                full = os.path.join(sendto_path, name)
                if not os.path.exists(full):
                    self._log(f"  {name}: файл не найден")
                    continue
                attrs = ctypes.windll.kernel32.GetFileAttributesW(full)
                if attrs == -1:
                    continue
                visible = value.strip().lower() == 'true'
                new_attrs = (attrs & ~0x2) if visible else (attrs | 0x2)
                ctypes.windll.kernel32.SetFileAttributesW(full, new_attrs)
                self._log(f"  {name}: {'виден' if visible else 'скрыт'}")
        except Exception as e:
            self._log(f"  ОШИБКА sendto: {e}")

    @staticmethod
    def _log(message, log_name="apply_settings.log"):
        try:
            path = os.path.join(os.path.expanduser("~"), "Desktop", log_name)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
            with open(path, "a", encoding="utf-8") as f:
                f.write(f"[{ts}] {message}\n")
        except Exception:
            pass

    def _check_modified_files(self):
        modified = []
        for fname, expected in MOD_EXPECTED_HASHES.items():
            path = os.path.join(os.environ["WINDIR"], "System32", fname)
            if not os.path.exists(path):
                continue
            try:
                with open(path, "rb") as f:
                    if hashlib.sha256(f.read()).hexdigest().upper() == expected.upper():
                        modified.append(fname)
            except Exception:
                pass
        return modified

    def _restore_modified_files(self):
        modified_names = []
        for fname in MOD_EXPECTED_HASHES:
            path = os.path.join(os.environ["WINDIR"], "System32", fname)
            if not os.path.exists(path):
                continue
            try:
                with open(path, "rb") as f:
                    h = hashlib.sha256(f.read()).hexdigest().upper()
                if h == MOD_EXPECTED_HASHES[fname].upper():
                    modified_names.append(fname)
            except Exception:
                pass

        if not modified_names:
            return

        try:
            subprocess.run("taskkill /f /im explorer.exe", shell=True,
                           capture_output=True, creationflags=CREATE_NO_WINDOW)
            time.sleep(0.5)
        except Exception:
            pass

        for fname in modified_names:
            path = os.path.join(os.environ["WINDIR"], "System32", fname)
            backup = f"{path}_orig"
            try:
                subprocess.run(f'takeown /f "{path}" /a >nul 2>&1',
                               shell=True, capture_output=True,
                               creationflags=CREATE_NO_WINDOW)
                subprocess.run(f'icacls "{path}" /grant "%USERNAME%":(F) >nul 2>&1',
                               shell=True, capture_output=True,
                               creationflags=CREATE_NO_WINDOW)
                os.remove(path)
                if os.path.exists(backup):
                    os.rename(backup, path)
                    subprocess.run(
                        f'icacls "{path}" /setowner "NT SERVICE\\TrustedInstaller" >nul 2>&1',
                        shell=True, capture_output=True,
                        creationflags=CREATE_NO_WINDOW)
            except Exception as e:
                self._log(f"Ошибка при восстановлении {fname}: {e}")

        try:
            subprocess.Popen(['explorer.exe'], creationflags=CREATE_NO_WINDOW)
        except Exception:
            pass

    @staticmethod
    def _reset_desktop_icons_direct():
        user_clsid = '{59031A47-3F72-44A7-89C5-5595FE6B30EE}'
        computer_clsid = '{20D04FE0-3AEA-1069-A2D8-08002B30309D}'
        network_clsid = '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}'

        defaults = (
            (f'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{user_clsid}\\DefaultIcon',
             r'C:\Windows\System32\imageres.dll,-123'),
            (f'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{computer_clsid}\\DefaultIcon',
             r'C:\Windows\System32\imageres.dll,-109'),
            (f'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{network_clsid}\\DefaultIcon',
             r'C:\Windows\System32\imageres.dll,-25'),
        )
        for path, icon in defaults:
            _reg('add', path, '/ve', '/d', icon, '/f')

        _reg('add', RECYCLE_PATH, '/ve',
             '/d', r'C:\Windows\System32\imageres.dll,-55', '/f')
        _reg('add', RECYCLE_PATH, '/v', 'Empty',
             '/d', r'C:\Windows\System32\imageres.dll,-55', '/f')
        _reg('add', RECYCLE_PATH, '/v', 'Full',
             '/d', r'C:\Windows\System32\imageres.dll,-54', '/f')

        cp_defaults = {
            '{26EE0668-A00A-44D7-9371-BEB064C98683}': r'C:\Windows\System32\imageres.dll,-21',
            '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}': r'C:\Windows\System32\imageres.dll,-27',
            '{21EC2020-3AEA-1069-A2DD-08002B30309D}': r'C:\Windows\System32\imageres.dll,-21',
        }
        for clsid, icon in cp_defaults.items():
            path = (f'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\'
                    f'{clsid}\\DefaultIcon')
            _reg('add', path, '/ve', '/d', icon, '/f')

        lib_value = r'C:\Windows\System32\imageres.dll,-1023'
        for hive, path in (
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
            (winreg.HKEY_CURRENT_USER,
             r"Software\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
            (winreg.HKEY_CLASSES_ROOT,
             r"CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
        ):
            try:
                try:
                    k = winreg.OpenKey(hive, path, 0, winreg.KEY_WRITE)
                except FileNotFoundError:
                    k = winreg.CreateKey(hive, path)
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, lib_value)
                winreg.CloseKey(k)
                break
            except (PermissionError, OSError):
                continue
            except Exception:
                continue

        subprocess.run(['ie4uinit.exe', '-show'], capture_output=True,
                       creationflags=CREATE_NO_WINDOW)

    def reset_all_settings(self):
        modified = self._check_modified_files()
        msg = ("ВНИМАНИЕ! Это действие сбросит все настройки к стандартным значениям Windows.\n\n"
               "Будут сброшены:\n• Иконки рабочего стола\n• Иконки дисков\n"
               "• Стрелки ярлыков\n• Текст '- ярлык'\n• Пункты контекстного меню\n"
               "• Настройки Проводника\n• Настройки Панели задач\n"
               "• Пункты меню 'Создать'\n• Пункты меню 'Отправить'\n")
        if modified:
            msg += "\n⚠️ БУДУТ ВОССТАНОВЛЕНЫ МОДИФИЦИРОВАННЫЕ ФАЙЛЫ:\n"
            for f in modified:
                msg += f"  • {f}\n"
        msg += "\n\nПродолжить?"

        if not messagebox.askyesno("Сброс всех настроек", msg, icon='warning'):
            return

        errors = []

        if modified:
            try:
                self._restore_modified_files()
            except Exception as e:
                errors.append(f"restore dll: {e}")

        try:
            self._reset_desktop_icons_direct()
        except Exception as e:
            errors.append(f"desktop: {e}")

        try:
            _reg('delete', DRIVE_ICONS, '/f')
        except Exception as e:
            errors.append(f"disk: {e}")

        try:
            from str import Module as StrModule
            temp = ttk.Frame(self.parent)
            s = StrModule(temp)
            temp.pack_forget()
            s.arrow_choice.set(1)
            s.shortcut_text_choice.set(1)
            s.apply_changes()
            s.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"str: {e}")

        try:
            from contextmenu import DesktopMenuTab
            temp = ttk.Frame(self.parent)
            dt = DesktopMenuTab(temp)
            dt.set_default_settings()
            temp.destroy()
        except Exception as e:
            errors.append(f"contextmenu: {e}")

        try:
            from explorer import ExplorerModule, IMAGE_DETAILS_KEY, IMAGE_DETAILS_VALUE
            temp = ttk.Frame(self.parent)
            e = ExplorerModule(temp)
            temp.pack_forget()
            is_admin = self._is_admin()

            if is_admin:
                for guid in e.checkbox_vars:
                    e.checkbox_vars[guid].set(True)
                e.apply_folder_changes()

            if is_admin and not e.bitlocker_enabled:
                e.toggle_bitlocker()

            if is_admin and e.photo_viewer_enabled:
                e.toggle_photo_viewer()

            if e.file_ext_enabled:
                e.toggle_file_ext()

            if not e.recent_folders_enabled:
                e.toggle_recent_folders()

            if getattr(e, 'image_details_supported', False):
                if getattr(e, 'image_details_hidden', False):
                    e.toggle_image_details()
            else:
                _reg('delete', IMAGE_DETAILS_KEY, '/v', IMAGE_DETAILS_VALUE, '/f')

            e.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"explorer: {e}")

        try:
            from taskbar import Module as TaskbarModule
            temp = ttk.Frame(self.parent)
            t = TaskbarModule(temp)
            temp.pack_forget()
            t._reset_all()
            t.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"taskbar: {e}")

        try:
            from createnew import (
                _reg_exists, _rename_key,
                TARGET_EXTENSIONS, MS_EXTENSIONS,
                _ext_registered, _find_shellnew_under_ext,
            )
            import winreg as _wr

            def _enable(hive, base):
                try:
                    if _reg_exists(hive, base + '_no'):
                        _rename_key(hive, base + '_no', base)
                except Exception:
                    pass

            _enable(_wr.HKEY_CLASSES_ROOT, r'Folder\ShellNew')
            for _, path in TARGET_EXTENSIONS.items():
                _enable(_wr.HKEY_LOCAL_MACHINE, path)

            for ext in MS_EXTENSIONS:
                if not _ext_registered(ext):
                    continue
                for hive in (_wr.HKEY_LOCAL_MACHINE, _wr.HKEY_CURRENT_USER):
                    for h, path, cls, enabled in _find_shellnew_under_ext(hive, ext):
                        if path.endswith('_no'):
                            try:
                                _rename_key(h, path, path[:-3])
                            except Exception:
                                pass
        except Exception as e:
            errors.append(f"createnew: {e}")

        try:
            import ctypes
            sendto_path = os.path.join(os.environ.get('APPDATA', ''),
                                       'Microsoft', 'Windows', 'SendTo')
            if os.path.isdir(sendto_path):
                for name in os.listdir(sendto_path):
                    if name.lower() == 'desktop.ini':
                        continue
                    full = os.path.join(sendto_path, name)
                    if not (os.path.isfile(full) or os.path.islink(full)):
                        continue
                    attrs = ctypes.windll.kernel32.GetFileAttributesW(full)
                    if attrs == -1:
                        continue
                    ctypes.windll.kernel32.SetFileAttributesW(full, attrs & ~0x2)
        except Exception as e:
            errors.append(f"sendto: {e}")

        self._restart_explorer_silent()

        if errors:
            messagebox.showwarning("Сброс завершён с ошибками", "\n".join(errors[:5]))

    @staticmethod
    def _restart_explorer_silent():
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                           capture_output=True, creationflags=CREATE_NO_WINDOW)
            time.sleep(0.5)
            subprocess.Popen(['explorer.exe'], creationflags=CREATE_NO_WINDOW)
        except Exception:
            pass

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.frame.pack_forget()