import ctypes
import os
import string
import subprocess
import tkinter as tk
from tkinter import ttk

import logger
from core import reg_utils
from core.ui_utils import add_tooltip, MessageLabel

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
MOD = "disk"

ICON_PACKAGES = {
    "Windows 95": {"Дискета": 101, "Дисковод": 102, "Жесткий диск": 103, "Флешка": 105},
    "Windows 98": {"Дискета": 101, "Дисковод": 102, "Жесткий диск": 103, "Флешка": 105},
    "Windows 2000": {"Дискета": 101, "Дисковод": 102, "Жесткий диск": 103, "Флешка": 105},
    "Windows XP": {"Дискета": 401, "Дисковод": 402, "Жесткий диск": 403,
                   "Жесткий диск (sys)": 404, "Флешка": 405},
    "Windows Vista": {"Дискета": 501, "Дисковод": 502, "Жесткий диск": 503,
                      "Жесткий диск (sys)": 504, "Флешка": 505},
    "Windows 7": {"Дискета": 501, "Дисковод": 502, "Жесткий диск": 503,
                  "Жесткий диск (sys)": 504, "Флешка": 505},
    "Windows 10": {"Дискета": 701, "Дисковод": 702, "Жесткий диск": 703,
                   "Жесткий диск (sys)": 704, "Флешка": 705},
    "Windows 11": {"Дискета": 801, "Дисковод": 802, "Жесткий диск": 803,
                   "Жесткий диск (sys)": 804, "Флешка": 805},
}

DRIVE_ICONS = r'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons'


class Module:
    def __init__(self, parent, module_name=None):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.module_name = module_name or "disk"
        self.dll_path = self._get_dll_path()
        self.msg = None
        self._is_destroyed = False
        self._create_widgets()
        self.show()

    def _get_dll_path(self):
        try:
            from main import DLLChecker
            return DLLChecker.get_dll_path_for_module(self.parent, self.module_name)
        except ImportError:
            return None

    @staticmethod
    def _get_available_disks():
        return [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]

    def clear_icon_cache(self):
        if self._is_destroyed:
            return
        try:
            self.msg.show("Очистка кэша иконок...", "info")
            subprocess.run(['ie4uinit.exe', '-ClearIconCache'],
                           capture_output=True, timeout=30,
                           creationflags=CREATE_NO_WINDOW)
            subprocess.run(['ie4uinit.exe', '-show'],
                           capture_output=True, timeout=30,
                           creationflags=CREATE_NO_WINDOW)
            logger.log("icon cache cleared", module=MOD)
            try:
                ctypes.windll.user32.SystemParametersInfoW(
                    0x0014, 0, None, 0x0002 | 0x0001)
            except Exception:
                pass
            self.msg.show("Кэш иконок успешно очищен и обновлен!", "success")
        except subprocess.TimeoutExpired:
            self.msg.show("Очистка кэша заняла слишком много времени", "warning")
        except Exception as e:
            logger.log(f"ERROR clear_icon_cache: {e}", module=MOD)
            self.msg.show(f"Ошибка при очистке кэша: {e}", "warning")

    def _create_widgets(self):
        ttk.Label(self.frame, text="Настройка иконок дисков",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)

        settings = ttk.LabelFrame(self.frame, text="Настройки иконок дисков", padding=15)
        settings.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # disk
        row = ttk.Frame(settings)
        row.pack(fill=tk.X, pady=8)
        ttk.Label(row, text="Выберите диск:", font=('Arial', 9), width=15).pack(
            side=tk.LEFT, padx=(0, 10))
        self.disk_var = tk.StringVar()
        self.disk_combobox = ttk.Combobox(row, textvariable=self.disk_var, width=20)
        self.disk_combobox['values'] = self._get_available_disks()
        self.disk_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if self.disk_combobox['values']:
            self.disk_combobox.current(0)

        # package
        row = ttk.Frame(settings)
        row.pack(fill=tk.X, pady=8)
        ttk.Label(row, text="Выберите пакет:", font=('Arial', 9), width=15).pack(
            side=tk.LEFT, padx=(0, 10))
        self.package_var = tk.StringVar()
        self.package_combobox = ttk.Combobox(row, textvariable=self.package_var, width=20)
        self.package_combobox['values'] = list(ICON_PACKAGES.keys())
        self.package_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.package_combobox.current(0)
        self.package_combobox.bind("<<ComboboxSelected>>", self._update_icon_combobox)

        # icon
        row = ttk.Frame(settings)
        row.pack(fill=tk.X, pady=8)
        ttk.Label(row, text="Выберите иконку:", font=('Arial', 9), width=15).pack(
            side=tk.LEFT, padx=(0, 10))
        self.icon_var = tk.StringVar()
        self.icon_combobox = ttk.Combobox(row, textvariable=self.icon_var, width=20)
        self.icon_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._update_icon_combobox()

        # apply
        row = ttk.Frame(settings)
        row.pack(fill=tk.X, pady=15)
        self.apply_button = ttk.Button(
            row, text="Применить иконку",
            command=self.apply_icon,
            state='normal' if (self.dll_path or logger.is_debug()) else 'disabled')
        self.apply_button.pack(fill=tk.X, pady=2)

        # delete + cache
        row = ttk.Frame(settings)
        row.pack(fill=tk.X, pady=5)
        ttk.Button(row, text="Удалить иконку",
                   command=self.remove_icon).pack(side=tk.LEFT, padx=(0, 5),
                                                   fill=tk.X, expand=True)
        ttk.Button(row, text="Удалить все иконки",
                   command=self.remove_all_icons).pack(side=tk.LEFT, padx=5,
                                                       fill=tk.X, expand=True)
        cache_btn = ttk.Button(row, text="Очистить кеш", command=self.clear_icon_cache)
        cache_btn.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        add_tooltip(cache_btn, "Очистить кеш иконок Windows")

        self.msg = MessageLabel(settings, wraplength=400)

    def _update_icon_combobox(self, event=None):
        icons = list(ICON_PACKAGES.get(self.package_var.get(), {}).keys())
        self.icon_combobox['values'] = icons
        if icons:
            self.icon_combobox.current(0)

    def apply_icon(self):
        if not self.dll_path and not logger.is_debug():
            return
        disk = self.disk_var.get().strip(":\\/").upper()
        icon_number = ICON_PACKAGES.get(self.package_var.get(), {}).get(self.icon_var.get())
        if not disk or not icon_number:
            return
        key = f'{DRIVE_ICONS}\\{disk}\\DefaultIcon'
        value = (f'{self.dll_path},-{icon_number}' if self.dll_path
                 else f'imaster.dll,-{icon_number}')
        rr = reg_utils.reg_add(key, None, value, 'REG_SZ', module=MOD)
        if rr.returncode == 0:
            self.msg.show(f"Иконка для диска {disk} установлена", "success")
        else:
            logger.log(f"reg_add failed for {key}: rc={rr.returncode}",
                       module=MOD)
            self.msg.show(f"Не удалось установить иконку для {disk}", "warning")

    def remove_icon(self):
        disk = self.disk_var.get().strip(":\\/").upper()
        if not disk:
            return
        key = f'{DRIVE_ICONS}\\{disk}'
        reg_utils.reg_delete(key, module=MOD)
        self.msg.show(f"Иконка для диска {disk} удалена", "success")

    def remove_all_icons(self):
        reg_utils.reg_delete(DRIVE_ICONS, module=MOD)
        self.msg.show("Все иконки дисков удалены", "success")

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self._is_destroyed = True
        self.frame.pack_forget()