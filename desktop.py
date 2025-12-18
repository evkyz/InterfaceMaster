import ctypes
import ctypes.wintypes
import subprocess
import sys
import tkinter as tk
import winreg
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

        self.dll_path = self.get_dll_path()
        self.message_timer = None

        self.initialize_registry_paths()
        self.initialize_icon_sets()

        self.create_widgets()
        self.show()

    def get_dll_path(self):
        if not HAS_DLL_CHECKER:
            return None

        dll_path = DLLChecker.get_dll_path_for_module(self.parent, self.module_name)
        return dll_path

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
            "Windows 95": {
                "user": "-1",
                "computer": "-2",
                "network": "-3",
                "recycle_bin_empty": "-4",
                "recycle_bin_full": "-5",
                "control_panel": "-6",
                "libraries": None,
            },
            "Windows 98": {
                "user": "-1",
                "computer": "-12",
                "network": "-13",
                "recycle_bin_empty": "-4",
                "recycle_bin_full": "-5",
                "control_panel": "-6",
                "libraries": None,
            },
            "Windows 2000": {
                "user": "-1",
                "computer": "-22",
                "network": "-13",
                "recycle_bin_empty": "-24",
                "recycle_bin_full": "-25",
                "control_panel": "-6",
                "libraries": None,
            },
            "Windows XP": {
                "user": "-31",
                "computer": "-32",
                "network": "-33",
                "recycle_bin_empty": "-34",
                "recycle_bin_full": "-35",
                "control_panel": "-36",
                "libraries": None,
            },
            "Windows Vista": {
                "user": "-41",
                "computer": "-42",
                "network": "-43",
                "recycle_bin_empty": "-44",
                "recycle_bin_full": "-45",
                "control_panel": "-46",
                "libraries": "-47",
            },
            "Windows 7": {
                "user": "-51",
                "computer": "-42",
                "network": "-53",
                "recycle_bin_empty": "-44",
                "recycle_bin_full": "-45",
                "control_panel": "-56",
                "libraries": "-47",
            },
            "Windows 10": {
                "user": "-61",
                "computer": "-62",
                "network": "-63",
                "recycle_bin_empty": "-64",
                "recycle_bin_full": "-65",
                "control_panel": "-66",
                "libraries": "-67",
            },
            "Windows 11": {
                "user": "-71",
                "computer": "-72",
                "network": "-73",
                "recycle_bin_empty": "-74",
                "recycle_bin_full": "-75",
                "control_panel": "-76",
                "libraries": "-77",
            },
            "New Year": {
                "user": "-81",
                "computer": "-82",
                "network": "-83",
                "recycle_bin_empty": "-84",
                "recycle_bin_full": "-85",
                "control_panel": "-86",
                "libraries": None,
            },
        }

        self.ICON_SETS = {}
        for name, icons in BASE_ICON_SETS.items():
            self.ICON_SETS[name] = icons.copy()

    def show_message(self, message, message_type="info"):
        if self.message_timer:
            self.parent.after_cancel(self.message_timer)

        self.message_label.config(text=message)

        if message_type == "info":
            self.message_label.config(foreground='#2c3e50')
        elif message_type == "success":
            self.message_label.config(foreground='#27ae60')
        elif message_type == "warning":
            self.message_label.config(foreground='#e74c3c')

        self.message_label.pack(pady=5)
        self.message_timer = self.parent.after(3000, self.clear_message)

    def clear_message(self):
        self.message_label.config(text="")
        self.message_label.pack_forget()
        self.message_timer = None

    def create_widgets(self):
        main_container = ttk.Frame(self.frame)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        title_label = ttk.Label(
            main_container,
            text="Настройка иконок Рабочего стола",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=(0, 15))

        desc_label = ttk.Label(
            main_container,
            text="Измените внешний вид иконок рабочего стола на стиль различных версий Windows",
            font=('Arial', 11),
            foreground='#34495e',
            justify=tk.CENTER,
            wraplength=400
        )
        desc_label.pack(pady=(0, 25))

        selection_frame = ttk.LabelFrame(main_container, text="Выбор набора иконок", padding=15)
        selection_frame.pack(fill=tk.X, pady=10)

        grid_frame = ttk.Frame(selection_frame)
        grid_frame.pack(fill=tk.X, expand=True)

        ttk.Label(grid_frame, text="Набор иконок:", font=('Arial', 10)).grid(
            row=0, column=0, padx=(0, 10), pady=5, sticky="w"
        )

        self.icon_set_var = tk.StringVar()
        self.icon_set_combo = ttk.Combobox(
            grid_frame,
            textvariable=self.icon_set_var,
            values=list(self.ICON_SETS.keys()),
            state="readonly",
            width=25,
            font=('Arial', 10)
        )
        self.icon_set_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.icon_set_combo.set("Windows 95")

        self.checkbox_var = tk.BooleanVar(value=False)
        self.checkbox = ttk.Checkbutton(
            grid_frame,
            text="",
            variable=self.checkbox_var,
            state="disabled"
        )
        self.checkbox.grid(row=0, column=2, padx=(20, 5), pady=5, sticky="w")

        self.icon_set_combo.bind("<<ComboboxSelected>>", self.on_icon_set_select)

        button_frame = ttk.Frame(main_container)
        button_frame.pack(pady=25, fill=tk.X)

        center_frame = ttk.Frame(button_frame)
        center_frame.pack()

        self.apply_btn = ttk.Button(
            center_frame,
            text="Применить",
            command=self.apply_icon_set,
            width=20,
        )
        self.apply_btn.pack(side=tk.LEFT, padx=5)

        self.reset_btn = ttk.Button(
            center_frame,
            text="Сбросить",
            command=self.reset_icons,
            width=20
        )
        self.reset_btn.pack(side=tk.LEFT, padx=5)

        self.message_label = ttk.Label(
            main_container,
            text="",
            font=('Arial', 9),
            wraplength=400
        )

        self.update_buttons_state()

    def update_buttons_state(self):
        if self.dll_path:
            self.apply_btn.config(state='normal')
        else:
            self.apply_btn.config(state='disabled')

        self.reset_btn.config(state='normal')

    def on_icon_set_select(self, event=None):
        selected_set = self.icon_set_var.get()
        if selected_set == "Windows XP":
            self.checkbox.config(state="normal")
        else:
            self.checkbox.config(state="disabled")
            self.checkbox_var.set(False)

    def create_tooltip(self, widget, text):
        def enter(event):
            x, y, cx, cy = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            self.tooltip_window = tk.Toplevel(widget)
            self.tooltip_window.wm_overrideredirect(True)
            self.tooltip_window.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip_window, text=text,
                             background="lightyellow", relief="solid",
                             borderwidth=1, font=("Arial", 8), padx=5, pady=2)
            label.pack()

        def leave(event):
            if hasattr(self, 'tooltip_window') and self.tooltip_window:
                self.tooltip_window.destroy()
                delattr(self, 'tooltip_window')

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def remove_tooltip(self, widget):
        widget.unbind("<Enter>")
        widget.unbind("<Leave>")
        if hasattr(self, 'tooltip_window') and self.tooltip_window:
            self.tooltip_window.destroy()
            delattr(self, 'tooltip_window')

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def request_admin(self):
        if not self.is_admin():
            response = messagebox.askyesno(
                "Требуются права администратора",
                "Для изменения иконок рабочего стола требуются права администратора.\n"
                "Перезапустить приложение с правами администратора?"
            )
            if response:
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, " ".join(sys.argv), None, 1
                )
                sys.exit()

    def refresh_icon_cache(self):
        try:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True, timeout=10)
            return True
        except:
            return False

    def set_icon(self, reg_path, icon_value, param_name="", use_local_machine=False, use_classes_root=False,
                 is_reset=False):
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
                    messagebox.showerror("Ошибка", "Внимание! DLL файл не найден!")
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

        for reg_path in control_panel_paths:
            try:
                if is_reset:
                    default_icons = {
                        self.CONTROL_PANEL_CLSIDS['main']: r"C:\Windows\System32\imageres.dll,-21",
                        self.CONTROL_PANEL_CLSIDS['category_view']: r"C:\Windows\System32\imageres.dll,-27",
                        self.CONTROL_PANEL_CLSIDS['all_items']: r"C:\Windows\System32\imageres.dll,-21"
                    }

                    for clsid, default_value in default_icons.items():
                        if clsid in reg_path:
                            icon_value = default_value
                            break

                self.set_icon(reg_path, icon_value, is_reset=is_reset)
                success_count += 1
            except Exception:
                pass

        return success_count

    def reset_libraries_icon(self):
        try:
            reg_path = self.REG_PATHS["libraries"]
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reg_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, r"C:\Windows\System32\imageres.dll,-1023")
            winreg.CloseKey(key)
            return True
        except:
            return False

    def apply_icon_set(self):
        if not self.dll_path:
            messagebox.showerror("Ошибка", "Внимание! DLL файл не найден!")
            return

        if not self.is_admin():
            self.request_admin()
            return

        selected_set = self.icon_set_var.get()
        if not selected_set:
            messagebox.showwarning("Внимание", "Пожалуйста, выберите набор иконок.")
            return

        icon_set = self.ICON_SETS.get(selected_set)
        if not icon_set:
            messagebox.showerror("Ошибка", "Неверный набор иконок.")
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
            self.show_message(message_text, "success")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при применении иконок:\n\n{str(e)}")

    def reset_icons(self):
        if not self.is_admin():
            self.request_admin()
            return

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
                        self.set_icon(reg_path, default_icons["recycle_bin_full"], "Full", is_reset=True)
                        self.set_icon(reg_path, default_icons["recycle_bin_full"], "", is_reset=True)
                    else:
                        self.set_icon(reg_path, value, is_reset=True)
                except Exception:
                    pass

            self.set_control_panel_icons("", is_reset=True)
            self.reset_libraries_icon()
            self.refresh_icon_cache()

            self.show_message("Иконки рабочего стола успешно сброшены на стандартные значения", "success")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при сбросе иконок:\n\n{str(e)}")

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.frame.pack_forget()