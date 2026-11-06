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


class Module:
    def __init__(self, parent, module_name=None):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.module_name = module_name or "desktop"

        self.REGISTRY_KEY_PATH = r"Software\InterfaceMaster"
        self.NEW_YEAR_OFFER_KEY = "NewYearOfferDeclined"
        self.NEW_YEAR_ICONS_INSTALLED = "NewYearIconsInstalled"
        self.REMOVE_OFFER_KEY = "RemoveNewYearOfferDeclined"

        self.dll_path = self.get_dll_path()

        self.message_timer = None
        self._is_destroyed = False
        self.tooltip_window = None

        self.initialize_registry_paths()
        self.initialize_icon_sets()
        self.initialize_custom_icon_vars()

        self.create_widgets()
        self.show()
        self.check_holiday_periods()

    def safe_clear_message(self):
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
                label = getattr(self, attr)
                if label.winfo_exists():
                    label.config(text="")
                    label.pack_forget()
            except Exception:
                pass

    def get_registry_value(self, key_name, default_value=None):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REGISTRY_KEY_PATH, 0,
                                winreg.KEY_READ) as key:
                value, _ = winreg.QueryValueEx(key, key_name)
                return value
        except FileNotFoundError:
            return default_value
        except Exception:
            return default_value

    def set_registry_value(self, key_name, value):
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.REGISTRY_KEY_PATH) as key:
                winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, str(value))
            return True
        except Exception:
            return False

    def is_new_year_offer_declined(self):
        return self.get_registry_value(self.NEW_YEAR_OFFER_KEY, "0") == "1"

    def mark_new_year_offer_declined(self):
        self.set_registry_value(self.NEW_YEAR_OFFER_KEY, "1")

    def is_remove_offer_declined(self):
        return self.get_registry_value(self.REMOVE_OFFER_KEY, "0") == "1"

    def mark_remove_offer_declined(self):
        self.set_registry_value(self.REMOVE_OFFER_KEY, "1")

    def are_new_year_icons_installed(self):
        if self.get_registry_value(self.NEW_YEAR_ICONS_INSTALLED, "0") == "1":
            return True
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REG_PATHS["user"], 0,
                                winreg.KEY_READ) as key:
                value, _ = winreg.QueryValueEx(key, "")
                if "imaster.dll" in value.lower() and ",-81" in value:
                    return True
        except Exception:
            pass
        return False

    def mark_new_year_icons_installed(self, installed=True):
        self.set_registry_value(self.NEW_YEAR_ICONS_INSTALLED, "1" if installed else "0")

    def get_dll_path(self):
        if not HAS_DLL_CHECKER:
            return None
        return DLLChecker.get_dll_path_for_module(self.parent, self.module_name)

    def get_dll_hash(self):
        try:
            from main import calculate_sha1
            if self.dll_path:
                return calculate_sha1(self.dll_path)
        except Exception:
            pass
        return None

    def is_new_hash(self):
        current_hash = self.get_dll_hash()
        if current_hash:
            return current_hash.upper() == 'A522AC997631927EA8CA441F1916E9375AD19805'
        return False

    def initialize_registry_paths(self):
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
            "libraries": r"CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon",
        }

    def initialize_icon_sets(self):
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

        self.ICON_SETS = {name: icons.copy() for name, icons in BASE_ICON_SETS.items()}
        if self.is_new_hash():
            self.ICON_SETS["New Year"]["libraries"] = "-87"

    def initialize_custom_icon_vars(self):
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
            "libraries": tk.StringVar(value="Windows Vista")
        }
        self.custom_icon_widgets = {}

    def check_holiday_periods(self):
        now = datetime.now()
        if now.month == 12 and 20 <= now.day <= 31:
            self.parent.after(500, self.check_new_year_offer)
        elif now.month == 1 and now.day == 14:
            self.parent.after(500, self.check_remove_offer)

    def check_new_year_offer(self):
        if self.is_new_year_offer_declined() or self.are_new_year_icons_installed():
            return
        self.show_new_year_offer()

    def show_new_year_offer(self):
        if not self.dll_path:
            return
        if messagebox.askyesno("Новогоднее предложение! 🎄",
                               "Хотите установить новогодние иконки на рабочий стол?"):
            self.apply_new_year_icons()
        else:
            self.mark_new_year_offer_declined()

    def check_remove_offer(self):
        if self.is_remove_offer_declined() or not self.are_new_year_icons_installed():
            return
        self.show_remove_offer()

    def show_remove_offer(self):
        if messagebox.askyesno("Праздники закончились! 🎉",
                               "Сбросить иконки на стандартные Windows?"):
            self.remove_new_year_icons()
        else:
            self.mark_remove_offer_declined()

    def apply_new_year_icons(self):
        if not self.dll_path:
            messagebox.showwarning("Внимание", "Файл imaster.dll не найден!")
            return
        self.icon_set_var.set("New Year")
        self.apply_icon_set()
        self.mark_new_year_icons_installed(True)

    def remove_new_year_icons(self):
        self.reset_icons()
        self.mark_new_year_icons_installed(False)
        self.show_message("Установлены стандартные иконки Windows", "success")

    def show_message(self, message, message_type="info", target="presets"):
        if self._is_destroyed:
            return
        if self.message_timer:
            try:
                self.parent.after_cancel(self.message_timer)
            except Exception:
                pass

        if target == "presets":
            label = getattr(self, 'presets_message_label', None)
        elif target == "custom":
            label = getattr(self, 'custom_message_label', None)
        else:
            return
        if label is None:
            return

        try:
            if label.winfo_exists():
                colors = {"info": '#2c3e50', "success": '#27ae60', "warning": '#e74c3c'}
                label.config(text=message, foreground=colors.get(message_type, '#2c3e50'))
                label.pack(pady=5)
                self.message_timer = self.parent.after(3000, self.safe_clear_message)
        except Exception:
            pass

    def clear_message(self):
        self.safe_clear_message()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.presets_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.presets_tab, text="Готовые наборы")

        self.custom_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.custom_tab, text="Индивидуальный выбор")

        self.create_presets_tab()
        self.create_custom_tab()

    def create_presets_tab(self):
        main_container = ttk.Frame(self.presets_tab)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        ttk.Label(main_container, text="Настройка иконок Рабочего стола",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=(0, 15))
        ttk.Label(main_container,
                  text="Измените внешний вид иконок рабочего стола на стиль различных версий Windows",
                  font=('Arial', 11), foreground='#34495e', justify=tk.CENTER,
                  wraplength=400).pack(pady=(0, 25))

        selection_frame = ttk.LabelFrame(main_container, text="Выбор набора иконок", padding=15)
        selection_frame.pack(fill=tk.X, pady=10)

        grid_frame = ttk.Frame(selection_frame)
        grid_frame.pack(fill=tk.X, expand=True)

        ttk.Label(grid_frame, text="Набор иконок:", font=('Arial', 10)).grid(
            row=0, column=0, padx=(0, 10), pady=5, sticky="w")

        preset_styles = ["Windows 95", "Windows 98", "Windows 2000", "Windows XP",
                         "Windows Vista", "Windows 7", "Windows 10", "Windows 11", "New Year"]
        self.icon_set_var = tk.StringVar()
        self.icon_set_combo = ttk.Combobox(grid_frame, textvariable=self.icon_set_var,
                                           values=preset_styles, state="readonly",
                                           width=25, font=('Arial', 10))
        self.icon_set_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.icon_set_combo.set("Windows 95")

        self.checkbox_var = tk.BooleanVar(value=False)
        self.checkbox = ttk.Checkbutton(grid_frame, text="", variable=self.checkbox_var,
                                        state="disabled")
        self.checkbox.grid(row=0, column=2, padx=(20, 5), pady=5, sticky="w")
        self.icon_set_combo.bind("<<ComboboxSelected>>", self.on_icon_set_select)

        button_frame = ttk.Frame(main_container)
        button_frame.pack(pady=25, fill=tk.X)
        center_frame = ttk.Frame(button_frame)
        center_frame.pack()

        self.presets_apply_btn = ttk.Button(center_frame, text="Применить",
                                            command=self.apply_icon_set, width=15)
        self.presets_apply_btn.pack(side=tk.LEFT, padx=5)
        self.presets_reset_btn = ttk.Button(center_frame, text="Сбросить",
                                            command=self.reset_icons, width=15)
        self.presets_reset_btn.pack(side=tk.LEFT, padx=5)
        self.presets_clear_cache_btn = ttk.Button(center_frame, text="Очистить кеш",
                                                  command=self.clear_icon_cache, width=15)
        self.presets_clear_cache_btn.pack(side=tk.LEFT, padx=5)
        self.create_tooltip(self.presets_clear_cache_btn, "Очистить кеш иконок Windows")

        self.presets_message_label = ttk.Label(main_container, text="", font=('Arial', 9),
                                               wraplength=400)
        self.update_buttons_state()

    def create_custom_tab(self):
        main_container = ttk.Frame(self.custom_tab)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        center_frame = ttk.Frame(main_container)
        center_frame.pack(anchor="n")

        ttk.Label(center_frame, text="Индивидуальный выбор иконок",
                  font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(pady=(0, 15))
        ttk.Label(center_frame,
                  text="Выберите стиль иконок \nдля каждого элемента рабочего стола",
                  font=('Arial', 11), foreground='#34495e',
                  justify=tk.CENTER).pack(pady=(0, 20))

        table_frame = ttk.Frame(center_frame)
        table_frame.pack()

        icon_names = {
            "user": "Пользователь",
            "computer": "Компьютер",
            "network": "Сеть",
            "recycle_bin_empty": "Корзина (пустая)",
            "recycle_bin_full": "Корзина (полная)",
            "control_panel": "Панель управления",
            "libraries": "Библиотеки"
        }

        for icon_key, display_name in icon_names.items():
            row_frame = ttk.Frame(table_frame)
            row_frame.pack(fill=tk.X, pady=5)
            ttk.Label(row_frame, text=display_name, width=20, anchor="e").pack(
                side=tk.LEFT, padx=(0, 10))
            if icon_key.startswith("recycle"):
                values = self.recycle_styles
            elif icon_key == "libraries":
                values = self.library_styles
            else:
                values = self.full_styles
            combo = ttk.Combobox(row_frame, textvariable=self.custom_icon_vars[icon_key],
                                 values=values, state="readonly", width=25)
            combo.pack(side=tk.LEFT, padx=5)
            self.custom_icon_widgets[icon_key] = combo

        button_frame = ttk.Frame(main_container)
        button_frame.pack(pady=25, fill=tk.X)
        center_button_frame = ttk.Frame(button_frame)
        center_button_frame.pack()

        self.custom_apply_btn = ttk.Button(center_button_frame, text="Применить",
                                           command=self.apply_custom_icons, width=15)
        self.custom_apply_btn.pack(side=tk.LEFT, padx=5)
        self.custom_reset_btn = ttk.Button(center_button_frame, text="Сбросить",
                                           command=self.reset_icons, width=15)
        self.custom_reset_btn.pack(side=tk.LEFT, padx=5)
        self.custom_clear_cache_btn = ttk.Button(center_button_frame, text="Очистить кеш",
                                                 command=self.clear_icon_cache, width=15)
        self.custom_clear_cache_btn.pack(side=tk.LEFT, padx=5)
        self.create_tooltip(self.custom_clear_cache_btn, "Очистить кеш иконок Windows")

        self.custom_message_label = ttk.Label(main_container, text="", font=('Arial', 9),
                                              wraplength=400)
        self.update_buttons_state()

    def update_buttons_state(self):
        if self._is_destroyed:
            return
        state = 'normal' if self.dll_path else 'disabled'
        for attr in ('presets_apply_btn', 'custom_apply_btn'):
            if hasattr(self, attr):
                getattr(self, attr).config(state=state)
        for attr in ('presets_reset_btn', 'custom_reset_btn',
                     'presets_clear_cache_btn', 'custom_clear_cache_btn'):
            if hasattr(self, attr):
                getattr(self, attr).config(state='normal')

    def on_icon_set_select(self, event=None):
        if self._is_destroyed:
            return
        if self.icon_set_var.get() == "Windows XP":
            self.checkbox.config(state="normal")
            self.create_tooltip(self.checkbox, "Включить иконки корзины от Whistler")
        else:
            self.checkbox.config(state="disabled")
            self.checkbox_var.set(False)
            self.remove_tooltip(self.checkbox)

    def create_tooltip(self, widget, text):
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
            if hasattr(self, 'tooltip_window') and self.tooltip_window:
                try:
                    self.tooltip_window.destroy()
                except Exception:
                    pass
                self.tooltip_window = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def remove_tooltip(self, widget):
        widget.unbind("<Enter>")
        widget.unbind("<Leave>")
        if self.tooltip_window:
            try:
                self.tooltip_window.destroy()
            except Exception:
                pass
            self.tooltip_window = None

    def refresh_icon_cache(self):
        try:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True, timeout=10)
            return True
        except Exception:
            return False

    def clear_icon_cache(self):
        if self._is_destroyed:
            return
        current_tab = self.notebook.index(self.notebook.select())
        target = "presets" if current_tab == 0 else "custom"
        try:
            self.show_message("Очистка кэша иконок...", "info", target)
            subprocess.run(['ie4uinit.exe', '-ClearIconCache'], capture_output=True,
                           text=True, timeout=30)
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True,
                           text=True, timeout=30)
            try:
                ctypes.windll.user32.SystemParametersInfoW(0x0014, 0, None, 0x0002 | 0x0001)
            except Exception:
                pass
            self.show_message("Кэш иконок успешно очищен и обновлен!", "success", target)
        except subprocess.TimeoutExpired:
            self.show_message("Очистка кэша заняла слишком много времени", "warning", target)
        except Exception as e:
            self.show_message(f"Ошибка при очистке кэша: {str(e)}", "warning", target)

    def set_icon(self, reg_path, icon_value, param_name="", use_local_machine=False,
                 use_classes_root=False, is_reset=False):
        try:
            if use_classes_root:
                registry_key = winreg.HKEY_CLASSES_ROOT
            elif use_local_machine:
                registry_key = winreg.HKEY_LOCAL_MACHINE
            else:
                registry_key = winreg.HKEY_CURRENT_USER
            try:
                key = winreg.OpenKey(registry_key, reg_path, 0, winreg.KEY_WRITE)
            except FileNotFoundError:
                key = winreg.CreateKey(registry_key, reg_path)

            if is_reset:
                full_value = icon_value
            else:
                if not self.dll_path:
                    return False
                full_value = f"{self.dll_path},{icon_value}"
            winreg.SetValueEx(key, param_name, 0, winreg.REG_SZ, full_value)
            winreg.CloseKey(key)
            return True
        except Exception:
            return False

    def set_control_panel_icons(self, icon_value, is_reset=False):
        success_count = 0
        control_panel_paths = [
            f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS['main']}\\DefaultIcon",
            f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS['category_view']}\\DefaultIcon",
            f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS['all_items']}\\DefaultIcon",
        ]
        default_icons = {
            self.CONTROL_PANEL_CLSIDS['main']: r"C:\Windows\System32\imageres.dll,-21",
            self.CONTROL_PANEL_CLSIDS['category_view']: r"C:\Windows\System32\imageres.dll,-27",
            self.CONTROL_PANEL_CLSIDS['all_items']: r"C:\Windows\System32\imageres.dll,-21"
        }
        for reg_path in control_panel_paths:
            try:
                if is_reset:
                    for clsid, default_value in default_icons.items():
                        if clsid in reg_path:
                            icon_value = default_value
                            break
                self.set_icon(reg_path, icon_value, is_reset=is_reset)
                success_count += 1
            except Exception:
                continue
        return success_count

    def reset_libraries_icon(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, self.REG_PATHS["libraries"], 0,
                                 winreg.KEY_WRITE)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                              r"C:\Windows\System32\imageres.dll,-1023")
            winreg.CloseKey(key)
            return True
        except Exception:
            return False

    def apply_icon_set(self):
        if self._is_destroyed or not self.dll_path:
            return
        selected_set = self.icon_set_var.get()
        icon_set = self.ICON_SETS.get(selected_set)
        if not icon_set:
            messagebox.showerror("Ошибка", "Неверный набор иконок")
            return

        try:
            for key, value in icon_set.items():
                if key == "libraries" and value is None:
                    self.reset_libraries_icon()
                    continue
                if selected_set == "Windows XP" and self.checkbox_var.get():
                    if key == "recycle_bin_empty":
                        value = "-37"
                    elif key == "recycle_bin_full":
                        value = "-38"

                if key == "recycle_bin_empty":
                    self.set_icon(self.REG_PATHS[key], value, "Empty")
                    self.set_icon(self.REG_PATHS[key], icon_set["recycle_bin_full"], "Full")
                    self.set_icon(self.REG_PATHS[key], icon_set["recycle_bin_full"], "")
                elif key == "libraries":
                    self.set_icon(self.REG_PATHS[key], value, "", use_classes_root=True)
                elif key == "control_panel":
                    self.set_control_panel_icons(value)
                else:
                    self.set_icon(self.REG_PATHS[key], value)

            self.refresh_icon_cache()

            message_text = f"Набор иконок '{selected_set}' успешно применён!"
            if selected_set == "Windows XP" and self.checkbox_var.get():
                message_text += " Применены иконки корзины от Whistler"

            if selected_set == "New Year":
                self.mark_new_year_icons_installed(True)
            elif self.are_new_year_icons_installed():
                self.mark_new_year_icons_installed(False)

            self.show_message(message_text, "success", "presets")
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Произошла ошибка при применении иконок:\n\n{str(e)}")

    def apply_custom_icons(self):
        if not self.dll_path:
            messagebox.showwarning("Внимание", "Файл imaster.dll не найден!")
            return
        try:
            applied_count = 0
            for icon_key, value_var in self.custom_icon_vars.items():
                value = self.get_icon_value(icon_key, value_var.get())
                if value is None:
                    continue
                if icon_key == "libraries":
                    self.set_icon(self.REG_PATHS[icon_key], value, "", use_classes_root=True)
                    applied_count += 1
                elif icon_key == "control_panel":
                    self.set_control_panel_icons(value)
                    applied_count += 1
                elif icon_key == "recycle_bin_empty":
                    self.set_icon(self.REG_PATHS[icon_key], value, "Empty")
                    applied_count += 1
                elif icon_key == "recycle_bin_full":
                    self.set_icon(self.REG_PATHS[icon_key], value, "Full")
                    self.set_icon(self.REG_PATHS[icon_key], value, "")
                    applied_count += 1
                else:
                    self.set_icon(self.REG_PATHS[icon_key], value)
                    applied_count += 1

            if applied_count > 0:
                self.refresh_icon_cache()
                self.show_message(f"Применено {applied_count} индивидуальных иконок",
                                  "success", "custom")
            else:
                self.show_message("Не выбрано ни одной иконки для применения",
                                  "warning", "custom")
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Произошла ошибка при применении иконок:\n\n{str(e)}")

    def get_icon_value(self, icon_key, style):
        if style == "Windows XP (Whistler)":
            return self.whistler_icons.get(icon_key)
        icon_set = self.ICON_SETS.get(style)
        if icon_set:
            return icon_set.get(icon_key)
        return None

    def reset_icons(self):
        if self._is_destroyed:
            return
        current_tab = self.notebook.index(self.notebook.select())
        target = "presets" if current_tab == 0 else "custom"
        try:
            default_icons = {
                "user": r"C:\Windows\System32\imageres.dll,-123",
                "computer": r"C:\Windows\System32\imageres.dll,-109",
                "network": r"C:\Windows\System32\imageres.dll,-25",
                "recycle_bin_empty": r"C:\Windows\System32\imageres.dll,-55",
                "recycle_bin_full": r"C:\Windows\System32\imageres.dll,-54",
            }
            for key, value in default_icons.items():
                try:
                    reg_path = self.REG_PATHS[key]
                    if key == "recycle_bin_empty":
                        self.set_icon(reg_path, value, "Empty", is_reset=True)
                        self.set_icon(reg_path, default_icons["recycle_bin_full"], "Full",
                                      is_reset=True)
                        self.set_icon(reg_path, default_icons["recycle_bin_full"], "",
                                      is_reset=True)
                    else:
                        self.set_icon(reg_path, value, is_reset=True)
                except Exception:
                    continue

            self.set_control_panel_icons("", is_reset=True)
            self.reset_libraries_icon()
            self.refresh_icon_cache()
            self.mark_new_year_icons_installed(False)

            for key in self.custom_icon_vars.keys():
                self.custom_icon_vars[key].set(
                    "Windows Vista" if key == "libraries" else "Windows 95")

            self.show_message(
                "Иконки рабочего стола успешно сброшены на стандартные значения",
                "success", target)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при сбросе иконок:\n\n{str(e)}")

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        try:
            self._is_destroyed = True
            self.safe_clear_message()
            self.frame.pack_forget()
        except Exception:
            pass

    def destroy(self):
        self.hide()