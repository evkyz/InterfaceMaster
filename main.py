import hashlib
import os
import shutil
import sys
import tkinter as tk
import winreg
from tkinter import ttk, messagebox

import platform
from PIL import Image, ImageTk


def calculate_sha1(filepath):
    """Вычисляет SHA-1 хеш файла"""
    try:
        sha1_hash = hashlib.sha1()
        with open(filepath, 'rb') as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha1_hash.update(byte_block)
        return sha1_hash.hexdigest().upper()
    except Exception:
        return None


def check_windows_version():
    """Проверяет версию Windows"""
    system = platform.system()
    release = platform.release()
    version = platform.version()

    if system != "Windows":
        return False, f"Не Windows система: {system}"

    try:
        version_parts = version.split('.')
        if len(version_parts) >= 2:
            major = int(version_parts[0])
            minor = int(version_parts[1])
            build = int(version_parts[2]) if len(version_parts) > 2 else 0

            if major == 10 and minor == 0 and build >= 22000:
                return False, f"Windows 11 (сборка {build})"
            elif major == 10 and minor == 0 and build < 22000:
                return True, "Windows 10"
            else:
                return False, f"Windows {major}.{minor} (сборка {build})"
        else:
            return False, f"Windows (неизвестная версия: {version})"
    except (ValueError, IndexError):
        return False, f"Windows (ошибка определения версии: {version})"


def show_warning_message():
    """Показывает предупреждение о системе"""
    warning_text = "Внимание! Корректная работа программы\nна данной системе не гарантируется!"
    messagebox.showwarning("Проверка системы", warning_text)


class DLLChecker:
    """Класс для централизованной проверки DLL"""

    dll_info = None
    copy_offered = False
    
    # Ключ реестра для сохранения пути к DLL
    REGISTRY_KEY_PATH = r"Software\InterfaceMaster"
    REGISTRY_DLL_PATH_KEY = "DLLPath"

    @staticmethod
    def save_dll_path_to_registry(dll_path):
        """Сохраняет путь к DLL в реестр"""
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, DLLChecker.REGISTRY_KEY_PATH) as key:
                winreg.SetValueEx(key, DLLChecker.REGISTRY_DLL_PATH_KEY, 0, winreg.REG_SZ, dll_path)
            return True
        except Exception:
            return False

    @staticmethod
    def get_dll_path_from_registry():
        """Получает путь к DLL из реестра"""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, DLLChecker.REGISTRY_KEY_PATH, 0, winreg.KEY_READ) as key:
                dll_path, _ = winreg.QueryValueEx(key, DLLChecker.REGISTRY_DLL_PATH_KEY)
                if os.path.exists(dll_path):
                    return dll_path
        except FileNotFoundError:
            pass
        except Exception:
            pass
        return None

    @staticmethod
    def initialize():
        """Инициализация проверки DLL"""
        if DLLChecker.dll_info is None:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            checker = InterfaceMaster.DLLCheckerInternal(base_path)
            DLLChecker.dll_info = checker.check_imaster_dll()

    @staticmethod
    def get_dll_info():
        """Получение информации о DLL"""
        if DLLChecker.dll_info is None:
            DLLChecker.initialize()
        return DLLChecker.dll_info.copy()

    @staticmethod
    def get_dll_path():
        """Получение пути к DLL если она найдена и хеш совпадает"""
        dll_info = DLLChecker.get_dll_info()
        if dll_info['dll_found'] and dll_info['hash_match']:
            return dll_info['dll_path']
        return None

    @staticmethod
    def copy_dll_to_system32(dll_path):
        """Копирует DLL файл в System32 (без лишних вопросов)"""
        try:
            if not dll_path or not os.path.exists(dll_path):
                return False, "Исходный файл DLL не найден"

            system32_path = r"C:\Windows\System32\imaster.dll"

            import ctypes

            # Просто копируем файл, перезаписывая существующий
            shutil.copy2(dll_path, system32_path)

            if os.path.exists(system32_path):
                DLLChecker.save_dll_path_to_registry(system32_path)
                DLLChecker.initialize()
                return True, "Файл успешно скопирован в System32"
            else:
                return False, "Не удалось скопировать файл"

        except PermissionError:
            return False, "Отказано в доступе. Попробуйте запустить программу от имени администратора."
        except Exception as e:
            return False, f"Ошибка при копировании: {str(e)}"

    @staticmethod
    def get_dll_status():
        """Получает полный статус DLL"""
        dll_info = DLLChecker.get_dll_info()

        system32_path = r"C:\Windows\System32\imaster.dll"
        system32_dll_exists = False
        system32_hash_match = False

        if os.path.exists(system32_path):
            system32_dll_exists = True
            system32_hash = calculate_sha1(system32_path)
            if system32_hash in InterfaceMaster.DLLCheckerInternal.VALID_HASHES:
                system32_hash_match = True

        saved_dll_path = DLLChecker.get_dll_path_from_registry()
        saved_path_valid = False
        if saved_dll_path and os.path.exists(saved_dll_path):
            saved_hash = calculate_sha1(saved_dll_path)
            if saved_hash in InterfaceMaster.DLLCheckerInternal.VALID_HASHES:
                saved_path_valid = True

        # Определяем рекомендуемый путь
        if saved_path_valid:
            recommended_path = saved_dll_path
            source = "saved_registry"
        elif system32_dll_exists and system32_hash_match:
            recommended_path = system32_path
            source = "system32"
        elif dll_info['dll_found'] and dll_info['hash_match']:
            recommended_path = dll_info['dll_path']
            source = "found"
        else:
            recommended_path = None
            source = "none"

        # Условие для предложения копирования/замены
        should_offer_copy = (
            dll_info['dll_found'] and
            dll_info['hash_match'] and
            (not system32_dll_exists or not system32_hash_match)
        )

        return {
            **dll_info,
            'system32_path': system32_path,
            'system32_exists': system32_dll_exists,
            'system32_hash_match': system32_hash_match,
            'recommended_path': recommended_path,
            'should_offer_copy': should_offer_copy,
            'is_in_system32': system32_dll_exists and system32_hash_match,
            'saved_path': saved_dll_path,
            'source': source,
            'saved_path_valid': saved_path_valid
        }

    @staticmethod
    def offer_copy_to_system32(parent_window, module_name):
        """
        Предлагает скопировать DLL в System32 для модулей desktop/disk.
        Возвращает путь к DLL, который должен использовать модуль.
        """
        dll_status = DLLChecker.get_dll_status()

        # Проверяем, нужно ли предлагать копирование
        if not dll_status['should_offer_copy']:
            return dll_status['recommended_path']

        # Предотвращаем множественные предложения
        if DLLChecker.copy_offered:
            return dll_status['recommended_path']

        DLLChecker.copy_offered = True

        # Формируем сообщение в зависимости от ситуации
        if dll_status['system32_exists'] and not dll_status['system32_hash_match']:
            message_text = (
                f"Файл imaster.dll в System32 повреждён!\n\n"
                f"Хотите заменить файл в C:\\Windows\\System32?\n\n"
            )
        else:
            message_text = (
                f"Хотите скопировать файл в C:\\Windows\\System32?\n\n"
            )

        response = messagebox.askyesno(
            "Копирование DLL",
            message_text
        )

        if response:
            success, message = DLLChecker.copy_dll_to_system32(dll_status['dll_path'])

            if success:
                if message:
                    messagebox.showinfo("Успех", message)
                DLLChecker.initialize()
                DLLChecker.save_dll_path_to_registry(dll_status['system32_path'])
                return dll_status['system32_path']
            else:
                if message:
                    messagebox.showwarning("Предупреждение",
                                           f"Не удалось скопировать файл:\n{message}\n\n"
                                           f"Будет использовано текущее расположение")
                DLLChecker.save_dll_path_to_registry(dll_status['dll_path'])
                return dll_status['dll_path']
        else:
            # Пользователь отказался - сохраняем текущий путь в реестр
            DLLChecker.save_dll_path_to_registry(dll_status['dll_path'])
            return dll_status['dll_path']

    @staticmethod
    def get_dll_path_for_module(parent_window, module_name):
        """
        Получает путь к DLL для конкретного модуля.
        Для модулей desktop и disk может предложить копирование.
        Для других модулей возвращает путь без предложений.
        """
        dll_status = DLLChecker.get_dll_status()

        if not dll_status['dll_found']:
            if parent_window and module_name in ["desktop", "disk"]:
                messagebox.showwarning("Внимание", "Файл imaster.dll не найден!")
            return None

        if not dll_status['hash_match']:
            if parent_window and module_name in ["desktop", "disk"]:
                messagebox.showwarning("Внимание", "Файл imaster.dll поврежден!")
            return None

        if module_name in ["desktop", "disk"]:
            return DLLChecker.offer_copy_to_system32(parent_window, module_name)
        else:
            return dll_status['recommended_path']

    @staticmethod
    def get_dll_path_without_prompt():
        """
        Получает путь к DLL без каких-либо запросов пользователю.
        Возвращает None если DLL не найдена или повреждена.
        """
        dll_status = DLLChecker.get_dll_status()

        if not dll_status['dll_found'] or not dll_status['hash_match']:
            return None

        return dll_status['recommended_path']


class InterfaceMaster:
    """Основной класс приложения"""

    class DLLCheckerInternal:
        """Внутренний класс для проверки DLL"""

        # Два валидных хеша SHA-1
        VALID_HASHES = [
            'F62A4326A8300D9B828744E24314B98433510195',  # Первый валидный хеш
            'A522AC997631927EA8CA441F1916E9375AD19805'  # Второй валидный хеш
        ]

        def __init__(self, base_path):
            self.base_path = base_path

        def search_dll_in_filesystem(self):
            """Поиск imaster.dll в файловой системе (только правильные версии)"""
            search_locations = [
                (r"C:\Windows\System32\imaster.dll", "System32"),
                (os.path.join(self.base_path, "imaster.dll"), "Папка exe"),
                (os.path.join(os.getcwd(), "imaster.dll"), "Рабочая папка"),
                (os.path.join(os.path.dirname(self.base_path), "imaster.dll"), "Родительская от exe"),
                (os.path.join(os.path.dirname(os.getcwd()), "imaster.dll"), "Родительская от рабочей"),
                (r"C:\Windows\SysWOW64\imaster.dll", "SysWOW64"),
                (r"C:\Windows\imaster.dll", "Windows"),
                (r"C:\imaster.dll", "Корень C"),
            ]

            for file_path, description in search_locations:
                if os.path.exists(file_path):
                    # Проверяем хеш
                    file_hash = calculate_sha1(file_path)
                    if file_hash and file_hash in self.VALID_HASHES:
                        return file_path

            return None

        def verify_dll_hash(self, dll_path):
            """Проверяет хеш DLL файла"""
            actual_hash = calculate_sha1(dll_path)
            hash_match = False

            if actual_hash:
                hash_match = (actual_hash in self.VALID_HASHES)

            return {
                'valid_hashes': self.VALID_HASHES,
                'actual_hash': actual_hash,
                'dll_path': dll_path,
                'hash_match': hash_match,
                'dll_found': True
            }

        def check_imaster_dll(self):
            """Проверяет imaster.dll во всех возможных расположениях"""
            # Сначала проверяем сохраненный путь из реестра (приоритет)
            saved_path = DLLChecker.get_dll_path_from_registry()
            if saved_path and os.path.exists(saved_path):
                result = self.verify_dll_hash(saved_path)
                if result['hash_match']:
                    return result
            
            # Ищем правильную DLL в файловой системе
            found_path = self.search_dll_in_filesystem()
            if found_path:
                return self.verify_dll_hash(found_path)

            # Если ничего не нашли
            return {
                'valid_hashes': self.VALID_HASHES,
                'actual_hash': None,
                'dll_path': None,
                'hash_match': False,
                'dll_found': False
            }

    def __init__(self, root):
        self.root = root
        self.root.title("Interface Master v2.1 Beta")
        self.root.geometry("520x710")
        self.root.resizable(False, False)

        self.windows_version_check()

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        self.base_path = base_path

        if base_path not in sys.path:
            sys.path.insert(0, base_path)

        icon_path = os.path.join(base_path, "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except:
                pass

        DLLChecker.initialize()

        self.version = "2.1.0.170 Beta"

        self.logo_image = None
        self.logo_photo = None

        self.modules_loaded = self.load_modules()
        self.current_module_instance = None

        self.center_window()
        self.setup_styles()
        self.load_logo()
        self.create_interface()

    def windows_version_check(self):
        """Проверяет версию Windows и показывает предупреждение при необходимости"""
        is_windows10, message = check_windows_version()

        if not is_windows10:
            self.root.after(100, show_warning_message)

    def load_modules(self):
        """Загружает все доступные модули с учетом PyInstaller"""
        modules_loaded = {}

        module_names = [
            'about', 'taskbar', 'disk', 'mod', 'str',
            'desktop', 'contextmenu', 'explorer'
        ]

        for module_name in module_names:
            try:
                if getattr(sys, 'frozen', False):
                    import importlib.util
                    module_filename = f"{module_name}.py"
                    module_path = os.path.join(self.base_path, module_filename)

                    if os.path.exists(module_path):
                        spec = importlib.util.spec_from_file_location(f"imod_{module_name}", module_path)
                        module = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(module)
                        modules_loaded[module_name] = module
                else:
                    module = __import__(module_name)
                    modules_loaded[module_name] = module
            except:
                continue

        return modules_loaded

    def load_logo(self):
        """Загрузка логотипа из файла"""
        logo_path = os.path.join(self.base_path, "logo.png")
        try:
            if os.path.exists(logo_path):
                image = Image.open(logo_path)
                image = image.resize((40, 40), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(image)
            else:
                self.logo_photo = None
        except:
            self.logo_photo = None

    def center_window(self):
        """Центрирование окна на экране"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_styles(self):
        """Настройка стилей для виджетов"""
        style = ttk.Style()
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        style.configure('Main.TButton', font=('Arial', 9), padding=10)
        style.configure('Back.TButton', font=('Arial', 8), padding=4, foreground='#34495e')
        style.configure('Help.TButton', font=('Arial', 10, 'bold'), padding=2, width=3)

    def create_interface(self):
        """Создание основного интерфейса"""
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        self.header_frame = ttk.Frame(self.main_container)
        self.header_frame.pack(fill=tk.X, pady=(10, 5))

        self.content_frame = ttk.Frame(self.main_container)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.footer_frame = ttk.Frame(self.main_container)
        self.footer_frame.pack(fill=tk.X, pady=(5, 10))

        self.create_header()
        self.create_menu()
        self.create_footer()

        self.module_frame = ttk.Frame(self.content_frame)

    def create_header(self):
        """Создание заголовка с логотипом и кнопкой помощи"""
        header_content = ttk.Frame(self.header_frame)
        header_content.pack(expand=True, fill=tk.X, padx=20)

        if self.logo_photo:
            logo_label = ttk.Label(header_content, image=self.logo_photo)
            logo_label.grid(row=0, column=0, sticky="w", padx=(0, 10))

        title_label = ttk.Label(
            header_content,
            text="Interface Master",
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=1)

        help_btn = ttk.Button(
            header_content,
            text="?",
            style='Help.TButton',
            command=self.open_about
        )
        help_btn.grid(row=0, column=2, sticky="e", padx=(10, 0))

        header_content.grid_columnconfigure(0, weight=1)
        header_content.grid_columnconfigure(1, weight=0)
        header_content.grid_columnconfigure(2, weight=1)

    def create_menu(self):
        """Создание главного меню с кнопками"""
        self.menu_frame = ttk.Frame(self.content_frame)
        self.menu_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        buttons_data = [
            ("Настройка иконок Рабочего стола", "desktop", self.open_desktop),
            ("Настройка иконок дисков", "disk", self.open_disk),
            ("Настройка стрелок ярлыков", "str", self.open_arrows),
            ("Настройка контекстного меню", "contextmenu", self.open_context_menu),
            ("Настройка Панели задач", "taskbar", self.open_taskbar),
            ("Проводник", "explorer", self.open_explorer),
            ("Модификации", "mod", self.open_modifications)
        ]

        for text, module_name, command in buttons_data:
            btn = ttk.Button(
                self.menu_frame,
                text=text,
                style='Main.TButton',
                command=command
            )
            btn.pack(fill=tk.X, pady=3, ipady=6)

    def create_footer(self):
        """Создание нижней части"""
        left_footer_frame = ttk.Frame(self.footer_frame)
        left_footer_frame.pack(side=tk.LEFT, padx=10)

        self.back_btn = ttk.Button(
            left_footer_frame,
            text="Назад",
            style='Back.TButton',
            command=self.back_to_menu,
            width=15
        )
        self.back_btn.grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.back_btn.grid_remove()

        self.version_label = ttk.Label(
            left_footer_frame,
            text=f"v{self.version}",
            font=('Arial', 7),
            foreground='#95a5a6'
        )
        self.version_label.grid(row=1, column=0, sticky="w")

        right_footer_frame = ttk.Frame(self.footer_frame)
        right_footer_frame.pack(side=tk.RIGHT, padx=10)

        self.copyright_label = ttk.Label(
            right_footer_frame,
            text="© EvKyz, 2026",
            font=('Arial', 7),
            foreground='#95a5a6'
        )
        self.copyright_label.pack(side=tk.RIGHT)

    def open_about(self):
        """Открыть модуль 'О программе'"""
        self.open_module("about")

    def open_module(self, module_name):
        """Открывает интерфейс модуля"""
        if module_name not in self.modules_loaded:
            messagebox.showerror("Ошибка", f"Модуль {module_name} не загружен")
            return

        self.menu_frame.pack_forget()

        self.module_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.back_btn.grid()

        for widget in self.module_frame.winfo_children():
            widget.destroy()

        try:
            if module_name == 'about':
                self.current_module_instance = self.modules_loaded[module_name].AboutModule(
                    parent=self.module_frame,
                    version=self.version,
                    dll_info=DLLChecker.get_dll_status()
                )
            elif module_name == 'explorer':
                self.current_module_instance = self.modules_loaded[module_name].ExplorerModule(
                    parent=self.module_frame
                )
            else:
                module_class = getattr(self.modules_loaded[module_name], 'Module', None)
                if module_class:
                    if module_name in ["desktop", "disk"]:
                        self.current_module_instance = module_class(
                            parent=self.module_frame,
                            module_name=module_name
                        )
                    else:
                        self.current_module_instance = module_class(
                            parent=self.module_frame
                        )
                else:
                    error_label = ttk.Label(
                        self.module_frame,
                        text=f"Модуль {module_name} не имеет правильной структуры",
                        font=('Arial', 10),
                        foreground='red',
                        justify=tk.CENTER
                    )
                    error_label.pack(expand=True)
                    return

        except Exception:
            error_label = ttk.Label(
                self.module_frame,
                text=f"Ошибка загрузки модуля {module_name}",
                font=('Arial', 10),
                foreground='red',
                justify=tk.CENTER
            )
            error_label.pack(expand=True)

    def back_to_menu(self):
        """Возврат в главное меню"""
        self.module_frame.pack_forget()

        for widget in self.module_frame.winfo_children():
            widget.destroy()

        self.current_module_instance = None

        self.back_btn.grid_remove()

        self.menu_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    def open_desktop(self):
        self.open_module("desktop")

    def open_disk(self):
        self.open_module("disk")

    def open_arrows(self):
        self.open_module("str")

    def open_context_menu(self):
        self.open_module("contextmenu")

    def open_taskbar(self):
        self.open_module("taskbar")

    def open_modifications(self):
        self.open_module("mod")

    def open_explorer(self):
        self.open_module("explorer")


def main():
    """Основная функция запуска приложения"""
    root = tk.Tk()
    app = InterfaceMaster(root)
    root.mainloop()


if __name__ == "__main__":
    main()