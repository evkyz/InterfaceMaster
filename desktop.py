import ctypes
import ctypes.wintypes
import subprocess
import tkinter as tk
import winreg
from datetime import datetime
from tkinter import ttk, messagebox

# Импортируем централизованную информацию о DLL из main
try:
    from main import DLLChecker

    HAS_DLL_CHECKER = True
except ImportError:
    HAS_DLL_CHECKER = False


class Module:
    def __init__(self, parent, module_name=None):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.module_name = module_name or "desktop"  # Сохраняем имя модуля

        # Ключ реестра для сохранения состояния
        self.REGISTRY_KEY_PATH = r"Software\InterfaceMaster"
        self.NEW_YEAR_OFFER_KEY = "NewYearOfferDeclined"
        self.NEW_YEAR_ICONS_INSTALLED = "NewYearIconsInstalled"
        self.REMOVE_OFFER_KEY = "RemoveNewYearOfferDeclined"

        # Получаем путь к DLL через централизованный механизм
        self.dll_path = self.get_dll_path()

        # Таймер для очистки сообщений
        self.message_timer = None
        self._is_destroyed = False  # Флаг для отслеживания уничтожения виджета

        # Инициализация переменных
        self.initialize_registry_paths()
        self.initialize_icon_sets()

        # Создание виджетов
        self.create_widgets()
        self.show()

        # Проверяем периоды для предложений
        self.check_holiday_periods()

    def safe_clear_message(self):
        """Безопасная очистка сообщения с проверкой существования виджета"""
        if self._is_destroyed:
            return

        if self.message_timer:
            try:
                self.parent.after_cancel(self.message_timer)
            except:
                pass
            self.message_timer = None

        try:
            if self.message_label.winfo_exists():
                self.message_label.config(text="")
                self.message_label.pack_forget()
        except:
            pass

    def get_registry_value(self, key_name, default_value=None):
        """Получает значение из реестра"""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.REGISTRY_KEY_PATH, 0, winreg.KEY_READ) as key:
                value, value_type = winreg.QueryValueEx(key, key_name)
                return value
        except FileNotFoundError:
            return default_value
        except:
            return default_value

    def set_registry_value(self, key_name, value):
        """Устанавливает значение в реестре"""
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.REGISTRY_KEY_PATH) as key:
                winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, str(value))
            return True
        except:
            return False

    def is_new_year_offer_declined(self):
        """Проверяет, было ли отклонено новогоднее предложение"""
        declined = self.get_registry_value(self.NEW_YEAR_OFFER_KEY, "0")
        return declined == "1"

    def mark_new_year_offer_declined(self):
        """Помечает, что новогоднее предложение было отклонено"""
        self.set_registry_value(self.NEW_YEAR_OFFER_KEY, "1")

    def is_remove_offer_declined(self):
        """Проверяет, было ли отклонено предложение убрать новогодние иконки"""
        declined = self.get_registry_value(self.REMOVE_OFFER_KEY, "0")
        return declined == "1"

    def mark_remove_offer_declined(self):
        """Помечает, что предложение убрать новогодние иконки было отклонено"""
        self.set_registry_value(self.REMOVE_OFFER_KEY, "1")

    def are_new_year_icons_installed(self):
        """Проверяет, установлены ли новогодние иконки в данный момент"""
        # Проверяем запись в реестре
        installed = self.get_registry_value(self.NEW_YEAR_ICONS_INSTALLED, "0")
        if installed == "1":
            return True

        # Дополнительная проверка: смотрим, какие иконки сейчас установлены
        try:
            # Проверяем одну из иконок (например, для пользователя)
            reg_path = self.REG_PATHS["user"]
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ) as key:
                value, _ = winreg.QueryValueEx(key, "")
                # Проверяем, содержит ли значение новогоднюю иконку (-81)
                if "imaster.dll" in value.lower() and ",-81" in value:
                    return True
        except:
            pass

        return False

    def mark_new_year_icons_installed(self, installed=True):
        """Помечает, что новогодние иконки установлены или удалены"""
        value = "1" if installed else "0"
        self.set_registry_value(self.NEW_YEAR_ICONS_INSTALLED, value)

    def get_dll_path(self):
        """Получение пути к DLL через централизованный механизм"""
        if not HAS_DLL_CHECKER:
            # Это маловероятно, но на всякий случай
            return None

        # Используем специальный метод, который сам покажет предупреждения при необходимости
        dll_path = DLLChecker.get_dll_path_for_module(self.parent, self.module_name)
        return dll_path

    def get_dll_hash(self):
        """Получает хеш текущей DLL файла"""
        try:
            from main import calculate_sha1  # Импортируем из main
            if self.dll_path:
                return calculate_sha1(self.dll_path)
        except:
            pass
        return None

    def is_new_hash(self):
        """Проверяет, используется ли новый хеш"""
        current_hash = self.get_dll_hash()
        if current_hash:
            # Новый хеш
            NEW_HASH = 'A522AC997631927EA8CA441F1916E9375AD19805'
            return current_hash.upper() == NEW_HASH
        return False

    def initialize_registry_paths(self):
        """Инициализация путей реестра"""
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
        """Инициализация наборов иконок"""
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
                "libraries": None,  # По умолчанию None, будет изменено если новый хеш
            },
        }

        self.ICON_SETS = {}
        for name, icons in BASE_ICON_SETS.items():
            self.ICON_SETS[name] = icons.copy()

        # Если используется новый хеш, обновляем иконку библиотек для New Year
        if self.is_new_hash():
            self.ICON_SETS["New Year"]["libraries"] = "-87"

    def check_holiday_periods(self):
        """Проверяет праздничные периоды и предлагает действия"""
        current_date = datetime.now()

        # Проверяем период с 20 декабря по 31 декабря - предложение установить
        if (current_date.month == 12 and current_date.day >= 20) and \
                (current_date.month == 12 and current_date.day <= 31):
            self.parent.after(500, self.check_new_year_offer)

        # Проверяем 14 января - предложение убрать (только в этот день)
        elif current_date.month == 1 and current_date.day == 14:
            self.parent.after(500, self.check_remove_offer)

    def check_new_year_offer(self):
        """Проверяет новогодний период и предлагает установить иконки"""
        # Проверяем, не было ли уже отклонено предложение
        if self.is_new_year_offer_declined():
            return

        # Проверяем, не установлены ли уже новогодние иконки
        if self.are_new_year_icons_installed():
            return

        self.show_new_year_offer()

    def show_new_year_offer(self):
        """Показывает предложение установить новогодние иконки"""
        if not self.dll_path:
            return  # Не показываем предложение если нет DLL

        response = messagebox.askyesno(
            "Новогоднее предложение! 🎄",
            "Хотите установить новогодние иконки на рабочий стол?"
        )

        if response:
            self.apply_new_year_icons()
        else:
            # Пользователь отказался - сохраняем это в реестре
            self.mark_new_year_offer_declined()

    def check_remove_offer(self):
        """Проверяет, нужно ли предложить убрать новогодние иконки"""
        # Проверяем, не было ли уже отклонено предложение
        if self.is_remove_offer_declined():
            return

        # Проверяем, установлены ли еще новогодние иконки
        if not self.are_new_year_icons_installed():
            return

        self.show_remove_offer()

    def show_remove_offer(self):
        """Показывает предложение убрать новогодние иконки"""
        response = messagebox.askyesno(
            "Праздники закончились! 🎉",
            "Сбросить иконки на стандартные Windows?"
        )

        if response:
            self.remove_new_year_icons()
        else:
            # Пользователь отказался - сохраняем это в реестре
            self.mark_remove_offer_declined()

    def apply_new_year_icons(self):
        """Автоматически применяет новогодние иконки"""
        if not self.dll_path:
            messagebox.showwarning("Внимание", "Файл imaster.dll не найден!")
            return

        # Устанавливаем набор New Year в комбобоксе
        self.icon_set_var.set("New Year")

        # Применяем иконки
        self.apply_icon_set()

        # Помечаем, что новогодние иконки установлены
        self.mark_new_year_icons_installed(True)

    def remove_new_year_icons(self):
        """Убирает новогодние иконки и возвращает стандартные через сброс"""
        # Выполняем сброс иконок на стандартные
        self.reset_icons()

        # Помечаем, что новогодние иконки убраны
        self.mark_new_year_icons_installed(False)

        # Показываем сообщение
        self.show_message("Установлены стандартные иконки Windows", "success")

    def show_message(self, message, message_type="info"):
        """Показывает сообщение под кнопками"""
        if self._is_destroyed:
            return

        if self.message_timer:
            try:
                self.parent.after_cancel(self.message_timer)
            except:
                pass

        if self.message_label.winfo_exists():
            self.message_label.config(text=message)

            if message_type == "info":
                self.message_label.config(foreground='#2c3e50')
            elif message_type == "success":
                self.message_label.config(foreground='#27ae60')
            elif message_type == "warning":
                self.message_label.config(foreground='#e74c3c')

            self.message_label.pack(pady=5)
            self.message_timer = self.parent.after(3000, self.safe_clear_message)

    def clear_message(self):
        """Очищает сообщение (старый метод для совместимости)"""
        self.safe_clear_message()

    def create_widgets(self):
        """Создание виджетов интерфейса"""
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
        """Обновляет состояние кнопок в зависимости от наличия DLL"""
        if self._is_destroyed:
            return

        if self.dll_path:
            self.apply_btn.config(state='normal')
        else:
            self.apply_btn.config(state='disabled')

        self.reset_btn.config(state='normal')

    def on_icon_set_select(self, event=None):
        """Обработчик выбора набора иконок"""
        if self._is_destroyed:
            return

        selected_set = self.icon_set_var.get()
        if selected_set == "Windows XP":
            self.checkbox.config(state="normal")
            self.create_tooltip(self.checkbox, "Включить иконки корзины от Whistler")
        else:
            self.checkbox.config(state="disabled")
            self.checkbox_var.set(False)
            self.remove_tooltip(self.checkbox)

    def create_tooltip(self, widget, text):
        """Создать всплывающую подсказку"""

        def enter(event):
            if self._is_destroyed:
                return
            try:
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
            except:
                pass

        def leave(event):
            if hasattr(self, 'tooltip_window') and self.tooltip_window:
                try:
                    self.tooltip_window.destroy()
                except:
                    pass
                delattr(self, 'tooltip_window')

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def remove_tooltip(self, widget):
        """Удалить привязки всплывающей подсказки"""
        widget.unbind("<Enter>")
        widget.unbind("<Leave>")
        if hasattr(self, 'tooltip_window') and self.tooltip_window:
            try:
                self.tooltip_window.destroy()
            except:
                pass
            delattr(self, 'tooltip_window')

    def is_admin(self):
        """Проверка прав администратора"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def refresh_icon_cache(self):
        """Обновление кэша иконок"""
        try:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True, timeout=10)
            return True
        except:
            return False

    def set_icon(self, reg_path, icon_value, param_name="", use_local_machine=False, use_classes_root=False,
                 is_reset=False):
        """Установка иконки в реестре"""
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
        """Установка иконок для панели управления во всех необходимых местах"""
        success_count = 0

        control_panel_paths = [
            f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS['main']}\\DefaultIcon",
            f"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS['category_view']}\\DefaultIcon",
            f"Software\\Microsoft\\Windows\CurrentVersion\\Explorer\\CLSID\\{self.CONTROL_PANEL_CLSIDS['all_items']}\\DefaultIcon",
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
            except:
                continue

        return success_count

    def reset_libraries_icon(self):
        """Сброс иконки библиотек на стандартную"""
        try:
            reg_path = self.REG_PATHS["libraries"]
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reg_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, r"C:\Windows\System32\imageres.dll,-1023")
            winreg.CloseKey(key)
            return True
        except:
            return False

    def apply_icon_set(self):
        """Применение выбранного набора иконок"""
        if self._is_destroyed:
            return

        if not self.dll_path:
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

            # Если устанавливается New Year, помечаем это
            if selected_set == "New Year":
                self.mark_new_year_icons_installed(True)
            elif selected_set != "New Year" and self.are_new_year_icons_installed():
                # Если была установлена New Year, а теперь другой набор - снимаем отметку
                self.mark_new_year_icons_installed(False)

            self.show_message(message_text, "success")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при применении иконок:\n\n{str(e)}")

    def reset_icons(self):
        """Сброс иконок на стандартные"""
        if self._is_destroyed:
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
                except:
                    continue

            self.set_control_panel_icons("", is_reset=True)
            self.reset_libraries_icon()
            self.refresh_icon_cache()

            # При сбросе иконок также снимаем отметку о новогодних иконках
            self.mark_new_year_icons_installed(False)

            self.show_message("Иконки рабочего стола успешно сброшены на стандартные значения", "success")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка при сбросе иконок:\n\n{str(e)}")

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        """Скрыть модуль"""
        try:
            self._is_destroyed = True
            self.safe_clear_message()
            self.frame.pack_forget()
        except:
            pass

    def destroy(self):
        """Полное уничтожение модуля"""
        try:
            self._is_destroyed = True
            if self.message_timer:
                try:
                    self.parent.after_cancel(self.message_timer)
                except:
                    pass
                self.message_timer = None

            if hasattr(self, 'tooltip_window') and self.tooltip_window:
                try:
                    self.tooltip_window.destroy()
                except:
                    pass
                delattr(self, 'tooltip_window')
        except:
            pass