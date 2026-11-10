import ctypes
import subprocess
import tkinter as tk
import winreg
from datetime import datetime
from tkinter import ttk, messagebox

try:
    from main import DLLChecker
    HAS_DLL_CHECKER = True
except ImportError:
    HAS_DLL_CHECKER = False

REG_ROOT = r"Software\InterfaceMaster"
NEW_YEAR_OFFER_KEY = "NewYearOfferDeclined"
NEW_YEAR_ICONS_INSTALLED = "NewYearIconsInstalled"
REMOVE_OFFER_KEY = "RemoveNewYearOfferDeclined"
NEW_YEAR_HASH = 'A522AC997631927EA8CA441F1916E9375AD19805'

BASE_ICON_SETS = {
    "Windows 95": {"user": "-1", "computer": "-2", "network": "-3",
                   "recycle_bin_empty": "-4", "recycle_bin_full": "-5",
                   "control_panel": "-6", "libraries": None},
    "Windows 98": {"user": "-1", "computer": "-12", "network": "-13",
                   "recycle_bin_empty": "-4", "recycle_bin_full": "-5",
                   "control_panel": "-6", "libraries": None},
    "Windows 2000": {"user": "-1", "computer": "-22", "network": "-13",
                     "recycle_bin_empty": "-24", "recycle_bin_full": "-25",
                     "control_panel": "-6", "libraries": None},
    "Windows XP": {"user": "-31", "computer": "-32", "network": "-33",
                   "recycle_bin_empty": "-34", "recycle_bin_full": "-35",
                   "control_panel": "-36", "libraries": None},
    "Windows Vista": {"user": "-41", "computer": "-42", "network": "-43",
                      "recycle_bin_empty": "-44", "recycle_bin_full": "-45",
                      "control_panel": "-46", "libraries": "-47"},
    "Windows 7": {"user": "-51", "computer": "-42", "network": "-53",
                  "recycle_bin_empty": "-44", "recycle_bin_full": "-45",
                  "control_panel": "-56", "libraries": "-47"},
    "Windows 10": {"user": "-61", "computer": "-62", "network": "-63",
                   "recycle_bin_empty": "-64", "recycle_bin_full": "-65",
                   "control_panel": "-66", "libraries": "-67"},
    "Windows 11": {"user": "-71", "computer": "-72", "network": "-73",
                   "recycle_bin_empty": "-74", "recycle_bin_full": "-75",
                   "control_panel": "-76", "libraries": "-77"},
    "New Year": {"user": "-81", "computer": "-82", "network": "-83",
                 "recycle_bin_empty": "-84", "recycle_bin_full": "-85",
                 "control_panel": "-86", "libraries": None},
}

DEFAULT_ICONS = {
    "user": r"C:\Windows\System32\imageres.dll,-123",
    "computer": r"C:\Windows\System32\imageres.dll,-109",
    "network": r"C:\Windows\System32\imageres.dll,-25",
    "recycle_bin_empty": r"C:\Windows\System32\imageres.dll,-55",
    "recycle_bin_full": r"C:\Windows\System32\imageres.dll,-54",
}

CONTROL_PANEL_DEFAULTS = {
    "{26EE0668-A00A-44D7-9371-BEB064C98683}": r"C:\Windows\System32\imageres.dll,-21",
    "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}": r"C:\Windows\System32\imageres.dll,-27",
    "{21EC2020-3AEA-1069-A2DD-08002B30309D}": r"C:\Windows\System32\imageres.dll,-21",
}

LIBRARIES_DEFAULT = r"C:\Windows\System32\imageres.dll,-1023"
LIBRARIES_CLSID = r"{031E4825-7B94-4dc3-B131-E946B44C8DD5}"


class Module:
    def __init__(self, parent, module_name=None):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.module_name = module_name or "desktop"

        self.dll_path = self._get_dll_path()
        self.message_timer = None
        self._is_destroyed = False
        self._init_registry_paths()
        self._init_icon_sets()
        self._init_custom_vars()
        self._create_widgets()
        self.show()
        self._check_holiday_periods()

    # ---- registry values ----
    def _get_reg_value(self, name, default=None):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_ROOT) as k:
                v, _ = winreg.QueryValueEx(k, name)
                return v
        except (FileNotFoundError, OSError):
            return default

    def _set_reg_value(self, name, value):
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_ROOT) as k:
                winreg.SetValueEx(k, name, 0, winreg.REG_SZ, str(value))
            return True
        except Exception:
            return False

    def _is_offer_declined(self, key):
        return self._get_reg_value(key, "0") == "1"

    def _mark_offer_declined(self, key):
        self._set_reg_value(key, "1")

    def _are_new_year_installed(self):
        if self._get_reg_value(NEW_YEAR_ICONS_INSTALLED, "0") == "1":
            return True
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                self.REG_PATHS["user"]) as k:
                value, _ = winreg.QueryValueEx(k, "")
                return "imaster.dll" in value.lower() and ",-81" in value
        except Exception:
            return False

    def _mark_new_year_installed(self, installed=True):
        self._set_reg_value(NEW_YEAR_ICONS_INSTALLED, "1" if installed else "0")

    # ---- dll ----
    def _get_dll_path(self):
        if not HAS_DLL_CHECKER:
            return None
        return DLLChecker.get_dll_path_for_module(self.parent, self.module_name)

    def _dll_hash(self):
        try:
            from main import calculate_sha1
            return calculate_sha1(self.dll_path) if self.dll_path else None
        except Exception:
            return None

    def _is_new_hash(self):
        h = self._dll_hash()
        return bool(h) and h.upper() == NEW_YEAR_HASH

    # ---- init ----
    def _init_registry_paths(self):
        self.CONTROL_PANEL_CLSIDS = {
            "main": "{26EE0668-A00A-44D7-9371-BEB064C98683}",
            "category_view": "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}",
            "all_items": "{21EC2020-3AEA-1069-A2DD-08002B30309D}",
        }
        self.REG_PATHS = {
            "user": r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon",
            "computer": r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon",
            "network": r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon",
            "recycle_bin_empty": r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon",
            "recycle_bin_full": r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon",
            "libraries": rf"CLSID\{LIBRARIES_CLSID}\DefaultIcon",
        }

    def _init_icon_sets(self):
        self.ICON_SETS = {name: icons.copy() for name, icons in BASE_ICON_SETS.items()}
        if self._is_new_hash():
            self.ICON_SETS["New Year"]["libraries"] = "-87"

    def _init_custom_vars(self):
        self.full_styles = list(self.ICON_SETS.keys())
        self.recycle_styles = []
        for style in self.full_styles:
            self.recycle_styles.append(style)
            if style == "Windows XP":
                self.recycle_styles.append("Windows XP (Whistler)")
        self.library_styles = [s for s in self.full_styles
                               if self.ICON_SETS[s].get("libraries") is not None]
        self.whistler_icons = {"recycle_bin_empty": "-37", "recycle_bin_full": "-38"}
        self.custom_icon_vars = {
            "user": tk.StringVar(value="Windows 95"),
            "computer": tk.StringVar(value="Windows 95"),
            "network": tk.StringVar(value="Windows 95"),
            "recycle_bin_empty": tk.StringVar(value="Windows 95"),
            "recycle_bin_full": tk.StringVar(value="Windows 95"),
            "control_panel": tk.StringVar(value="Windows 95"),
            "libraries": tk.StringVar(value="Windows Vista"),
        }

    # ---- holiday offers ----
    def _check_holiday_periods(self):
        now = datetime.now()
        if now.month == 12 and 20 <= now.day <= 31:
            self.parent.after(500, self._check_new_year_offer)
        elif now.month == 1 and now.day == 14:
            self.parent.after(500, self._check_remove_offer)

    def _check_new_year_offer(self):
        if not self.dll_path or self._is_offer_declined(NEW_YEAR_OFFER_KEY) \
                or self._are_new_year_installed():
            return
        if messagebox.askyesno("Новогоднее предложение! 🎄",
                               "Хотите установить новогодние иконки на рабочий стол?"):
            self.icon_set_var.set("New Year")
            self.apply_icon_set()
            self._mark_new_year_installed(True)
        else:
            self._mark_offer_declined(NEW_YEAR_OFFER_KEY)

    def _check_remove_offer(self):
        if self._is_offer_declined(REMOVE_OFFER_KEY) or not self._are_new_year_installed():
            return
        if messagebox.askyesno("Праздники закончились! 🎉",
                               "Сбросить иконки на стандартные Windows?"):
            self.reset_icons()
            self._mark_new_year_installed(False)
            self.show_message("Установлены стандартные иконки Windows", "success")
        else:
            self._mark_offer_declined(REMOVE_OFFER_KEY)

    # ---- messages ----
    def _safe_clear(self):
        if self._is_destroyed:
            return
        if self.message_timer:
            try:
                self.parent.after_cancel(self.message_timer)
            except Exception:
                pass
            self.message_timer = None
        for attr in ('presets_message_label', 'custom_message_label'):
            try:
                lbl = getattr(self, attr, None)
                if lbl and lbl.winfo_exists():
                    lbl.config(text="")
                    lbl.pack_forget()
            except Exception:
                pass

    def show_message(self, message, message_type="info", target="presets"):
        if self._is_destroyed:
            return
        if self.message_timer:
            try:
                self.parent.after_cancel(self.message_timer)
            except Exception:
                pass
        label = getattr(self, f'{target}_message_label', None)
        if label is None:
            return
        try:
            if label.winfo_exists():
                colors = {"info": '#2c3e50', "success": '#27ae60', "warning": '#e74c3c'}
                label.config(text=message, foreground=colors.get(message_type, '#2c3e50'))
                label.pack(pady=5)
                self.message_timer = self.parent.after(3000, self._safe_clear)
        except Exception:
            pass

    # ---- widgets ----
    def _create_widgets(self):
        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.presets_tab = ttk.Frame(self.notebook)
        self.custom_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.presets_tab, text="Готовые наборы")
        self.notebook.add(self.custom_tab, text="Индивидуальный выбор")
        self._create_presets_tab()
        self._create_custom_tab()

    def _create_presets_tab(self):
        main = ttk.Frame(self.presets_tab)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        ttk.Label(main, text="Настройка иконок Рабочего стола",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=(0, 15))
        ttk.Label(main,
                  text="Измените внешний вид иконок рабочего стола на стиль различных версий Windows",
                  font=('Arial', 11), foreground='#34495e', justify=tk.CENTER,
                  wraplength=400).pack(pady=(0, 25))

        sel = ttk.LabelFrame(main, text="Выбор набора иконок", padding=15)
        sel.pack(fill=tk.X, pady=10)
        grid = ttk.Frame(sel)
        grid.pack(fill=tk.X, expand=True)

        ttk.Label(grid, text="Набор иконок:", font=('Arial', 10)).grid(
            row=0, column=0, padx=(0, 10), pady=5, sticky="w")

        preset_styles = ["Windows 95", "Windows 98", "Windows 2000", "Windows XP",
                         "Windows Vista", "Windows 7", "Windows 10", "Windows 11", "New Year"]
        self.icon_set_var = tk.StringVar()
        self.icon_set_combo = ttk.Combobox(grid, textvariable=self.icon_set_var,
                                           values=preset_styles, state="readonly",
                                           width=25, font=('Arial', 10))
        self.icon_set_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.icon_set_combo.set("Windows 95")

        self.checkbox_var = tk.BooleanVar(value=False)
        self.checkbox = ttk.Checkbutton(grid, text="", variable=self.checkbox_var,
                                        state="disabled")
        self.checkbox.grid(row=0, column=2, padx=(20, 5), pady=5, sticky="w")
        self.icon_set_combo.bind("<<ComboboxSelected>>", self._on_icon_set_select)

        btns = ttk.Frame(main)
        btns.pack(pady=25, fill=tk.X)
        center = ttk.Frame(btns)
        center.pack()

        self.presets_apply_btn = ttk.Button(center, text="Применить",
                                            command=self.apply_icon_set, width=15)
        self.presets_apply_btn.pack(side=tk.LEFT, padx=5)
        self.presets_reset_btn = ttk.Button(center, text="Сбросить",
                                            command=self.reset_icons, width=15)
        self.presets_reset_btn.pack(side=tk.LEFT, padx=5)
        self.presets_clear_cache_btn = ttk.Button(center, text="Очистить кеш",
                                                  command=self.clear_icon_cache, width=15)
        self.presets_clear_cache_btn.pack(side=tk.LEFT, padx=5)
        self._create_tooltip(self.presets_clear_cache_btn, "Очистить кеш иконок Windows")

        self.presets_message_label = ttk.Label(main, text="", font=('Arial', 9),
                                               wraplength=400)
        self._update_buttons_state()

    def _create_custom_tab(self):
        main = ttk.Frame(self.custom_tab)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        center = ttk.Frame(main)
        center.pack(anchor="n")

        ttk.Label(center, text="Индивидуальный выбор иконок",
                  font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(pady=(0, 15))
        ttk.Label(center, text="Выберите стиль иконок\nдля каждого элемента рабочего стола",
                  font=('Arial', 11), foreground='#34495e',
                  justify=tk.CENTER).pack(pady=(0, 20))

        table = ttk.Frame(center)
        table.pack()

        names = {
            "user": "Пользователь", "computer": "Компьютер", "network": "Сеть",
            "recycle_bin_empty": "Корзина (пустая)",
            "recycle_bin_full": "Корзина (полная)",
            "control_panel": "Панель управления", "libraries": "Библиотеки",
        }
        for key, name in names.items():
            row = ttk.Frame(table)
            row.pack(fill=tk.X, pady=5)
            ttk.Label(row, text=name, width=20, anchor="e").pack(side=tk.LEFT, padx=(0, 10))
            if key.startswith("recycle"):
                values = self.recycle_styles
            elif key == "libraries":
                values = self.library_styles
            else:
                values = self.full_styles
            ttk.Combobox(row, textvariable=self.custom_icon_vars[key],
                         values=values, state="readonly", width=25).pack(side=tk.LEFT, padx=5)

        btns = ttk.Frame(main)
        btns.pack(pady=25, fill=tk.X)
        center_btns = ttk.Frame(btns)
        center_btns.pack()

        self.custom_apply_btn = ttk.Button(center_btns, text="Применить",
                                           command=self.apply_custom_icons, width=15)
        self.custom_apply_btn.pack(side=tk.LEFT, padx=5)
        self.custom_reset_btn = ttk.Button(center_btns, text="Сбросить",
                                           command=self.reset_icons, width=15)
        self.custom_reset_btn.pack(side=tk.LEFT, padx=5)
        self.custom_clear_cache_btn = ttk.Button(center_btns, text="Очистить кеш",
                                                 command=self.clear_icon_cache, width=15)
        self.custom_clear_cache_btn.pack(side=tk.LEFT, padx=5)
        self._create_tooltip(self.custom_clear_cache_btn, "Очистить кеш иконок Windows")

        self.custom_message_label = ttk.Label(main, text="", font=('Arial', 9),
                                              wraplength=400)
        self._update_buttons_state()

    def _update_buttons_state(self):
        if self._is_destroyed:
            return
        state = 'normal' if self.dll_path else 'disabled'
        for attr in ('presets_apply_btn', 'custom_apply_btn'):
            w = getattr(self, attr, None)
            if w:
                w.config(state=state)
        for attr in ('presets_reset_btn', 'custom_reset_btn',
                     'presets_clear_cache_btn', 'custom_clear_cache_btn'):
            w = getattr(self, attr, None)
            if w:
                w.config(state='normal')

    def _on_icon_set_select(self, event=None):
        if self._is_destroyed:
            return
        if self.icon_set_var.get() == "Windows XP":
            self.checkbox.config(state="normal")
            self._create_tooltip(self.checkbox, "Включить иконки корзины от Whistler")
        else:
            self.checkbox.config(state="disabled")
            self.checkbox_var.set(False)
            self._remove_tooltip(self.checkbox)

    def _create_tooltip(self, widget, text):
        self._remove_tooltip(widget)
        state = {'tip': None}

        def enter(event):
            if self._is_destroyed or state['tip']:
                return
            try:
                x = widget.winfo_rootx() + 25
                y = widget.winfo_rooty() + 25
                tip = tk.Toplevel(widget)
                tip.wm_overrideredirect(True)
                tip.wm_geometry(f"+{x}+{y}")
                tk.Label(tip, text=text, background="lightyellow",
                         relief="solid", borderwidth=1, font=("Arial", 8),
                         padx=5, pady=2).pack()
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

        widget._tooltip_state = state
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def _remove_tooltip(self, widget):
        state = getattr(widget, '_tooltip_state', None)
        if state and state.get('tip'):
            try:
                state['tip'].destroy()
            except Exception:
                pass
            state['tip'] = None
        try:
            widget.unbind("<Enter>")
            widget.unbind("<Leave>")
        except Exception:
            pass
        widget._tooltip_state = None

    # ---- cache / icons ----
    def _refresh_icon_cache(self):
        try:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True, timeout=10)
            return True
        except Exception:
            return False

    def clear_icon_cache(self):
        if self._is_destroyed:
            return
        target = "presets" if self.notebook.index(self.notebook.select()) == 0 else "custom"
        try:
            self.show_message("Очистка кэша иконок...", "info", target)
            subprocess.run(['ie4uinit.exe', '-ClearIconCache'],
                           capture_output=True, timeout=30)
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True, timeout=30)
            try:
                ctypes.windll.user32.SystemParametersInfoW(0x0014, 0, None, 0x0002 | 0x0001)
            except Exception:
                pass
            self.show_message("Кэш иконок успешно очищен и обновлен!", "success", target)
        except subprocess.TimeoutExpired:
            self.show_message("Очистка кэша заняла слишком много времени", "warning", target)
        except Exception as e:
            self.show_message(f"Ошибка при очистке кэша: {e}", "warning", target)

    # ---- icon write helpers ----

    def _write_icon_value(self, hive, path, param_name, value):
        """Записать значение иконки в указанный hive/path. Возвращает True/False."""
        try:
            try:
                key = winreg.OpenKey(hive, path, 0, winreg.KEY_WRITE)
            except FileNotFoundError:
                key = winreg.CreateKey(hive, path)
            winreg.SetValueEx(key, param_name, 0, winreg.REG_SZ, value)
            winreg.CloseKey(key)
            return True
        except (PermissionError, OSError):
            return False
        except Exception:
            return False

    def _set_icon(self, reg_path, icon_value, param_name="",
                  use_classes_root=False, is_reset=False):
        """Установка иконки.

        Для use_classes_root=True перебираем HKLM\\SOFTWARE\\Classes →
        HKCU\\Software\\Classes → HKCR. HKCR — виртуальный hive.
        """
        value = icon_value if is_reset else f"{self.dll_path},{icon_value}"

        if use_classes_root:
            candidates = [
                (winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Classes\\" + reg_path),
                (winreg.HKEY_CURRENT_USER, "Software\\Classes\\" + reg_path),
                (winreg.HKEY_CLASSES_ROOT, reg_path),
            ]
        else:
            candidates = [(winreg.HKEY_CURRENT_USER, reg_path)]

        for hive, path in candidates:
            if self._write_icon_value(hive, path, param_name, value):
                return True
        return False

    def _set_control_panel_icons(self, icon_value, is_reset=False):
        count = 0
        paths = [
            f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS[k]}\\DefaultIcon"
            for k in ("main", "category_view", "all_items")
        ]
        for reg_path in paths:
            try:
                current_value = icon_value
                if is_reset:
                    for clsid, default_value in CONTROL_PANEL_DEFAULTS.items():
                        if clsid in reg_path:
                            current_value = default_value
                            break
                self._set_icon(reg_path, current_value, is_reset=is_reset)
                count += 1
            except Exception:
                continue
        return count

    def _reset_libraries_icon(self):
        """Сброс иконки библиотек. HKCR виртуальный — пишем явно в HKLM → HKCU → HKCR."""
        for hive, path in (
            (winreg.HKEY_LOCAL_MACHINE,
             "SOFTWARE\\Classes\\CLSID\\" + LIBRARIES_CLSID + "\\DefaultIcon"),
            (winreg.HKEY_CURRENT_USER,
             "Software\\Classes\\CLSID\\" + LIBRARIES_CLSID + "\\DefaultIcon"),
            (winreg.HKEY_CLASSES_ROOT,
             "CLSID\\" + LIBRARIES_CLSID + "\\DefaultIcon"),
        ):
            if self._write_icon_value(hive, path, "", LIBRARIES_DEFAULT):
                return True
        return False

    def apply_icon_set(self):
        if self._is_destroyed or not self.dll_path:
            return
        selected = self.icon_set_var.get()
        icon_set = self.ICON_SETS.get(selected)
        if not icon_set:
            messagebox.showerror("Ошибка", "Неверный набор иконок")
            return
        try:
            for key, value in icon_set.items():
                if key == "libraries" and value is None:
                    self._reset_libraries_icon()
                    continue
                if selected == "Windows XP" and self.checkbox_var.get():
                    if key == "recycle_bin_empty":
                        value = "-37"
                    elif key == "recycle_bin_full":
                        value = "-38"

                if key == "recycle_bin_empty":
                    self._set_icon(self.REG_PATHS[key], value, "Empty")
                    self._set_icon(self.REG_PATHS[key], icon_set["recycle_bin_full"], "Full")
                    self._set_icon(self.REG_PATHS[key], icon_set["recycle_bin_full"], "")
                elif key == "libraries":
                    self._set_icon(self.REG_PATHS[key], value, "", use_classes_root=True)
                elif key == "control_panel":
                    self._set_control_panel_icons(value)
                else:
                    self._set_icon(self.REG_PATHS[key], value)

            self._refresh_icon_cache()

            text = f"Набор иконок '{selected}' успешно применён!"
            if selected == "Windows XP" and self.checkbox_var.get():
                text += " Применены иконки корзины от Whistler"

            if selected == "New Year":
                self._mark_new_year_installed(True)
            elif self._are_new_year_installed():
                self._mark_new_year_installed(False)

            self.show_message(text, "success", "presets")
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Произошла ошибка при применении иконок:\n\n{e}")

    def apply_custom_icons(self):
        if not self.dll_path:
            messagebox.showwarning("Внимание", "Файл imaster.dll не найден!")
            return
        try:
            count = 0
            for key, var in self.custom_icon_vars.items():
                value = self._get_icon_value(key, var.get())
                if value is None:
                    continue
                if key == "libraries":
                    self._set_icon(self.REG_PATHS[key], value, "", use_classes_root=True)
                elif key == "control_panel":
                    self._set_control_panel_icons(value)
                elif key == "recycle_bin_empty":
                    self._set_icon(self.REG_PATHS[key], value, "Empty")
                elif key == "recycle_bin_full":
                    self._set_icon(self.REG_PATHS[key], value, "Full")
                    self._set_icon(self.REG_PATHS[key], value, "")
                else:
                    self._set_icon(self.REG_PATHS[key], value)
                count += 1

            if count:
                self._refresh_icon_cache()
                self.show_message(f"Применено {count} индивидуальных иконок",
                                  "success", "custom")
            else:
                self.show_message("Не выбрано ни одной иконки для применения",
                                  "warning", "custom")
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Произошла ошибка при применении иконок:\n\n{e}")

    def _get_icon_value(self, key, style):
        if style == "Windows XP (Whistler)":
            return self.whistler_icons.get(key)
        icon_set = self.ICON_SETS.get(style)
        return icon_set.get(key) if icon_set else None

    def reset_icons(self):
        if self._is_destroyed:
            return
        target = "presets" if self.notebook.index(self.notebook.select()) == 0 else "custom"
        try:
            for key, value in DEFAULT_ICONS.items():
                try:
                    if key == "recycle_bin_empty":
                        self._set_icon(self.REG_PATHS[key], value, "Empty", is_reset=True)
                        self._set_icon(self.REG_PATHS[key], DEFAULT_ICONS["recycle_bin_full"],
                                       "Full", is_reset=True)
                        self._set_icon(self.REG_PATHS[key], DEFAULT_ICONS["recycle_bin_full"],
                                       "", is_reset=True)
                    else:
                        self._set_icon(self.REG_PATHS[key], value, is_reset=True)
                except Exception:
                    continue

            self._set_control_panel_icons("", is_reset=True)
            self._reset_libraries_icon()
            self._refresh_icon_cache()
            self._mark_new_year_installed(False)

            for key in self.custom_icon_vars:
                self.custom_icon_vars[key].set(
                    "Windows Vista" if key == "libraries" else "Windows 95")

            self.show_message("Иконки рабочего стола успешно сброшены на стандартные значения",
                              "success", target)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при сбросе иконок:\n\n{e}")

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        try:
            self._is_destroyed = True
            self._safe_clear()
            self.frame.pack_forget()
        except Exception:
            pass

    def destroy(self):
        self.hide()