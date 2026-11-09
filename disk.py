import os
import string
import subprocess
import sys
import tkinter as tk
from tkinter import ttk
import ctypes

try:
    from main import DLLChecker
    HAS_DLL_CHECKER = True
except ImportError:
    HAS_DLL_CHECKER = False

# Скрыть консольное окно при вызове reg.exe и т.п.
CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


def _run_hidden(args, **kwargs):
    """subprocess.run с гарантированно скрытым окном."""
    kwargs.setdefault("creationflags", CREATE_NO_WINDOW)
    return subprocess.run(args, **kwargs)


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
        self.message_timer = None
        self._is_destroyed = False
        self.tooltip_window = None
        self._create_widgets()
        self.show()

    def _get_dll_path(self):
        if not HAS_DLL_CHECKER:
            return None
        return DLLChecker.get_dll_path_for_module(self.parent, self.module_name)

    def show_message(self, message, message_type="info"):
        if self._is_destroyed:
            return
        if self.message_timer:
            try:
                self.parent.after_cancel(self.message_timer)
            except Exception:
                pass
        try:
            if self.message_label.winfo_exists():
                colors = {"info": '#2c3e50', "success": '#27ae60', "warning": '#e74c3c'}
                self.message_label.config(text=message,
                                          foreground=colors.get(message_type, '#2c3e50'))
                self.message_label.pack(pady=5)
                self.message_timer = self.parent.after(3000, self._clear_message)
        except Exception:
            pass

    def _clear_message(self):
        if self._is_destroyed:
            return
        try:
            if self.message_label.winfo_exists():
                self.message_label.config(text="")
                self.message_label.pack_forget()
        except Exception:
            pass
        self.message_timer = None

    @staticmethod
    def _get_available_disks():
        return [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]

    def _create_tooltip(self, widget, text):
        def enter(event):
            if self._is_destroyed:
                return
            try:
                x = widget.winfo_rootx() + 25
                y = widget.winfo_rooty() + 25
                self.tooltip_window = tk.Toplevel(widget)
                self.tooltip_window.wm_overrideredirect(True)
                self.tooltip_window.wm_geometry(f"+{x}+{y}")
                tk.Label(self.tooltip_window, text=text, background="lightyellow",
                         relief="solid", borderwidth=1, font=("Arial", 8),
                         padx=5, pady=2).pack()
            except Exception:
                pass

        def leave(event):
            if self.tooltip_window:
                try:
                    self.tooltip_window.destroy()
                except Exception:
                    pass
                self.tooltip_window = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def clear_icon_cache(self):
        if self._is_destroyed:
            return
        try:
            self.show_message("Очистка кэша иконок...", "info")
            _run_hidden(['ie4uinit.exe', '-ClearIconCache'],
                        capture_output=True, timeout=30)
            _run_hidden(['ie4uinit.exe', '-show'],
                        capture_output=True, timeout=30)
            try:
                ctypes.windll.user32.SystemParametersInfoW(0x0014, 0, None, 0x0002 | 0x0001)
            except Exception:
                pass
            self.show_message("Кэш иконок успешно очищен и обновлен!", "success")
        except subprocess.TimeoutExpired:
            self.show_message("Очистка кэша заняла слишком много времени", "warning")
        except Exception as e:
            self.show_message(f"Ошибка при очистке кэша: {e}", "warning")

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
        self.apply_button = ttk.Button(row, text="Применить иконку",
                                       command=self.apply_icon,
                                       state='normal' if self.dll_path else 'disabled')
        self.apply_button.pack(fill=tk.X, pady=2)

        # delete
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
        self._create_tooltip(cache_btn, "Очистить кеш иконок Windows")

        self.message_label = ttk.Label(settings, text="", font=('Arial', 9),
                                       wraplength=400)

    def _update_icon_combobox(self, event=None):
        icons = list(ICON_PACKAGES.get(self.package_var.get(), {}).keys())
        self.icon_combobox['values'] = icons
        if icons:
            self.icon_combobox.current(0)

    def apply_icon(self):
        if not self.dll_path:
            return
        disk = self.disk_var.get().strip(":\\/").upper()
        icon_number = ICON_PACKAGES.get(self.package_var.get(), {}).get(self.icon_var.get())
        if not disk or not icon_number:
            return
        try:
            _run_hidden(
                ['reg', 'add', f'{DRIVE_ICONS}\\{disk}\\DefaultIcon',
                 '/ve', '/d', f'{self.dll_path},-{icon_number}',
                 '/t', 'REG_SZ', '/f'],
                check=True, capture_output=True, text=True, encoding='cp866')
            self.show_message(f"Иконка для диска {disk} установлена", "success")
        except subprocess.CalledProcessError:
            pass

    def remove_icon(self):
        disk = self.disk_var.get().strip(":\\/").upper()
        if not disk:
            return
        try:
            _run_hidden(['reg', 'delete', f'{DRIVE_ICONS}\\{disk}', '/f'],
                        check=True, capture_output=True, text=True, encoding='cp866')
            self.show_message(f"Иконка для диска {disk} удалена", "success")
        except subprocess.CalledProcessError:
            pass

    def remove_all_icons(self):
        try:
            _run_hidden(['reg', 'delete', DRIVE_ICONS, '/f'], check=True,
                        capture_output=True, text=True, encoding='cp866')
            self.show_message("Все иконки дисков удалены", "success")
        except subprocess.CalledProcessError:
            pass

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self._is_destroyed = True
        self.frame.pack_forget()