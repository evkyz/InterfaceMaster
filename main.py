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
        """Копирует DLL файл в System32"""
        try:
            if not dll_path or not os.path.exists(dll_path):
                return False, "Исходный файл DLL не найден"

            system32_path = r"C:\Windows\System32\imaster.dll"

            import ctypes

            if os.path.exists(system32_path):
                existing_hash = calculate_sha1(system32_path)
                source_hash = calculate_sha1(dll_path)

                if existing_hash == source_hash:
                    return True, None  # Возвращаем None вместо сообщения
                else:
                    root = tk.Tk()
                    root.withdraw()

                    response = messagebox.askyesno(
                        "Перезапись файла",
                        "Файл imaster.dll уже существует в System32 с другим хешем\n"
                        "Хотите перезаписать его?"
                    )
                    root.destroy()

                    if not response:
                        return False, None

            shutil.copy2(dll_path, system32_path)

            if os.path.exists(system32_path):
                DLLChecker.initialize()
                return True, "Файл успешно скопирован в System32"
            else:
                return False, "Не удалось скопировать файл"

        except PermissionError:
            return False, "Отказано в доступе"
        except Exception as e:
            return False, f"Ошибка при копировании: {str(e)}"

    @staticmethod
    def check_registry_for_dll_path():
        """
        Проверяет реестр на наличие пути к imaster.dll
        Возвращает путь, если он найден в реестре, иначе None
        """
        registry_paths = [
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon"),
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon"),
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon"),
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon"),
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{26EE0668-A00A-44D7-9371-BEB064C98683}\DefaultIcon"),

            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"),

            (winreg.HKEY_CLASSES_ROOT, r"DesktopBackground\Shell\Reboot"),
            (winreg.HKEY_CLASSES_ROOT, r"DesktopBackground\Shell\Hibernation"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\DesktopBackground\Shell\Reboot"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\DesktopBackground\Shell\Hibernation"),
            (winreg.HKEY_CURRENT_USER, r"Software\Classes\DesktopBackground\Shell\Reboot"),
            (winreg.HKEY_CURRENT_USER, r"Software\Classes\DesktopBackground\Shell\Hibernation"),
        ]

        for hive, path in registry_paths:
            try:
                with winreg.OpenKey(hive, path) as key:
                    if "DriveIcons" in path:
                        try:
                            for i in range(winreg.QueryInfoKey(key)[0]):
                                drive_letter = winreg.EnumKey(key, i)
                                try:
                                    subkey_path = f"{path}\\{drive_letter}\\DefaultIcon"
                                    with winreg.OpenKey(hive, subkey_path) as subkey:
                                        value, _ = winreg.QueryValueEx(subkey, "")
                                        if "imaster.dll" in value.lower():
                                            dll_path = value.split(",")[0]
                                            if os.path.exists(dll_path):
                                                return dll_path
                                except:
                                    continue
                        except:
                            continue
                    else:
                        try:
                            value, _ = winreg.QueryValueEx(key, "")
                            if "imaster.dll" in value.lower():
                                dll_path = value.split(",")[0]
                                if os.path.exists(dll_path):
                                    return dll_path
                        except:
                            try:
                                value, _ = winreg.QueryValueEx(key, "Icon")
                                if "imaster.dll" in value.lower():
                                    dll_path = value.split(",")[0]
                                    if os.path.exists(dll_path):
                                        return dll_path
                            except:
                                continue
            except:
                continue

        return None

    @staticmethod
    def is_dll_in_registry():
        """Проверяет, используется ли imaster.dll в реестре"""
        return DLLChecker.check_registry_for_dll_path() is not None

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
            if system32_hash == dll_info['expected_sha1']:
                system32_hash_match = True

        registry_dll_path = DLLChecker.check_registry_for_dll_path()
        dll_in_registry = registry_dll_path is not None

        if dll_in_registry and registry_dll_path:
            recommended_path = registry_dll_path
            source = "registry"
        elif system32_dll_exists and system32_hash_match:
            recommended_path = system32_path
            source = "system32"
        elif dll_info['dll_found'] and dll_info['hash_match']:
            recommended_path = dll_info['dll_path']
            source = "found"
        else:
            recommended_path = None
            source = "none"

        should_offer_copy = (
                dll_info['dll_found'] and
                dll_info['hash_match'] and
                not (system32_dll_exists and system32_hash_match) and
                not dll_in_registry
        )

        return {
            **dll_info,
            'system32_path': system32_path,
            'system32_exists': system32_dll_exists,
            'system32_hash_match': system32_hash_match,
            'recommended_path': recommended_path,
            'should_offer_copy': should_offer_copy,
            'is_in_system32': system32_dll_exists and system32_hash_match,
            'is_in_registry': dll_in_registry,
            'registry_path': registry_dll_path,
            'source': source
        }

    @staticmethod
    def offer_copy_to_system32(parent_window, module_name):
        """
        Предлагает скопировать DLL в System32 для модулей desktop/disk.
        Возвращает путь к DLL, который должен использовать модуль.
        """
        dll_status = DLLChecker.get_dll_status()

        if DLLChecker.copy_offered:
            return dll_status['recommended_path']

        if dll_status['is_in_registry'] or dll_status['is_in_system32']:
            return dll_status['recommended_path']

        if not dll_status['should_offer_copy']:
            return dll_status['recommended_path']

        DLLChecker.copy_offered = True

        response = messagebox.askyesno(
            "Копирование DLL",
            f"Рекомендуется скопировать файл imaster.dll в System32\n\n"
            f"Текущее расположение: {dll_status['dll_path']}\n\n"
            f"Хотите скопировать файл в C:\\Windows\\System32?\n\n"
        )

        if response:
            success, message = DLLChecker.copy_dll_to_system32(dll_status['dll_path'])

            if success:
                # Показываем сообщение об успехе только если оно есть
                if message:
                    messagebox.showinfo("Успех", message)
                DLLChecker.initialize()
                return dll_status['system32_path']
            else:
                # Показываем предупреждение только если есть сообщение об ошибке
                if message:
                    messagebox.showwarning("Предупреждение",
                                           f"Не удалось скопировать файл:\n{message}\n\n"
                                           f"Будет использовано текущее расположение")
                return dll_status['dll_path']
        else:
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
                messagebox.showwarning("Внимание",
                                       f"Файл imaster.dll поврежден!\n\n"
                                       f"Ожидаемый хеш: {dll_status['expected_sha1']}\n"
                                       f"Фактический хеш: {dll_status['actual_hash']}")
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

        if dll_status['is_in_registry']:
            return dll_status['registry_path']

        if dll_status['is_in_system32']:
            return dll_status['system32_path']

        return dll_status['recommended_path']


class InterfaceMaster:
    """Основной класс приложения"""

    class DLLCheckerInternal:
        """Внутренний класс для проверки DLL"""

        EXPECTED_SHA1 = 'F62A4326A8300D9B828744E24314B98433510195'

        def __init__(self, base_path):
            self.base_path = base_path

        def search_dll_in_registry(self):
            """Поиск imaster.dll в реестре Windows"""
            dll_paths = set()

            try:
                key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        drive_letter = winreg.EnumKey(key, i)
                        try:
                            subkey_path = f"{key_path}\\{drive_letter}\\DefaultIcon"
                            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey_path) as subkey:
                                value, _ = winreg.QueryValueEx(subkey, "")
                                if "imaster.dll" in value.lower():
                                    dll_paths.add(value.split(",")[0])
                        except:
                            continue
            except:
                pass

            clsid_paths = [
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon",
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon",
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon",
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon",
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{26EE0668-A00A-44D7-9371-BEB064C98683}\DefaultIcon"
            ]

            for clsid_path in clsid_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_path) as key:
                        value, _ = winreg.QueryValueEx(key, "")
                        if "imaster.dll" in value.lower():
                            dll_paths.add(value.split(",")[0])
                except:
                    continue

            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"DesktopBackground\Shell\Reboot") as key:
                    try:
                        value_default, _ = winreg.QueryValueEx(key, "")
                        if "imaster.dll" in value_default.lower():
                            dll_paths.add(value_default.split(",")[0])
                    except:
                        pass

                    try:
                        value_icon, _ = winreg.QueryValueEx(key, "Icon")
                        if "imaster.dll" in value_icon.lower():
                            dll_paths.add(value_icon.split(",")[0])
                    except:
                        pass
            except:
                pass

            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"DesktopBackground\Shell\Hibernation") as key:
                    try:
                        value_default, _ = winreg.QueryValueEx(key, "")
                        if "imaster.dll" in value_default.lower():
                            dll_paths.add(value_default.split(",")[0])
                    except:
                        pass

                    try:
                        value_icon, _ = winreg.QueryValueEx(key, "Icon")
                        if "imaster.dll" in value_icon.lower():
                            dll_paths.add(value_icon.split(",")[0])
                    except:
                        pass
            except:
                pass

            additional_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\DesktopBackground\Shell\Reboot"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\DesktopBackground\Shell\Hibernation"),
                (winreg.HKEY_CURRENT_USER, r"Software\Classes\DesktopBackground\Shell\Reboot"),
                (winreg.HKEY_CURRENT_USER, r"Software\Classes\DesktopBackground\Shell\Hibernation"),
            ]

            for hive, path in additional_paths:
                try:
                    with winreg.OpenKey(hive, path) as key:
                        try:
                            value_default, _ = winreg.QueryValueEx(key, "")
                            if value_default and "imaster.dll" in value_default.lower():
                                dll_paths.add(value_default.split(",")[0])
                        except:
                            pass

                        try:
                            value_icon, _ = winreg.QueryValueEx(key, "Icon")
                            if value_icon and "imaster.dll" in value_icon.lower():
                                dll_paths.add(value_icon.split(",")[0])
                        except:
                            pass
                except:
                    continue

            return list(dll_paths)

        def verify_dll_hash(self, dll_path, expected_sha1):
            """Проверяет хеш DLL файла"""
            actual_hash = calculate_sha1(dll_path)
            hash_match = False

            if actual_hash:
                hash_match = (actual_hash == expected_sha1)

            return {
                'expected_sha1': expected_sha1,
                'actual_hash': actual_hash,
                'dll_path': dll_path,
                'hash_match': hash_match,
                'dll_found': True
            }

        def check_imaster_dll(self):
            """Проверяет imaster.dll во всех возможных расположениях"""
            search_locations = []

            system32_path = r"C:\Windows\System32\imaster.dll"
            search_locations.append((system32_path, "System32"))

            current_exe_dir = self.base_path
            exe_dir_path = os.path.join(current_exe_dir, "imaster.dll")
            search_locations.append((exe_dir_path, "Папка exe (base_path)"))

            work_dir = os.getcwd()
            work_dir_path = os.path.join(work_dir, "imaster.dll")
            search_locations.append((work_dir_path, "Рабочая папка"))

            if current_exe_dir:
                parent_dir = os.path.dirname(current_exe_dir)
                parent_dir_path = os.path.join(parent_dir, "imaster.dll")
                search_locations.append((parent_dir_path, "Родительская от base_path"))

            parent_work_dir = os.path.dirname(work_dir)
            parent_work_dir_path = os.path.join(parent_work_dir, "imaster.dll")
            search_locations.append((parent_work_dir_path, "Родительская от рабочей папки"))

            syswow64_path = r"C:\Windows\SysWOW64\imaster.dll"
            search_locations.append((syswow64_path, "SysWOW64"))

            windows_path = r"C:\Windows\imaster.dll"
            search_locations.append((windows_path, "Каталог Windows"))

            root_path = r"C:\imaster.dll"
            search_locations.append((root_path, "Корень диска C"))

            for file_path, description in search_locations:
                if os.path.exists(file_path):
                    return self.verify_dll_hash(file_path, self.EXPECTED_SHA1)

            registry_paths = self.search_dll_in_registry()

            if registry_paths:
                for reg_path in registry_paths:
                    if os.path.exists(reg_path):
                        return self.verify_dll_hash(reg_path, self.EXPECTED_SHA1)
                    else:
                        filename = os.path.basename(reg_path)
                        if filename.lower() == "imaster.dll":
                            for file_path, description in search_locations:
                                if os.path.exists(file_path):
                                    return self.verify_dll_hash(file_path, self.EXPECTED_SHA1)

            return {
                'expected_sha1': self.EXPECTED_SHA1,
                'actual_hash': None,
                'dll_path': None,
                'hash_match': False,
                'dll_found': False,
                'registry_paths': registry_paths if 'registry_paths' in locals() else []
            }

    def __init__(self, root):
        self.root = root
        self.root.title("Interface Master v1.1")
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

        self.version = "1.1.0.120"

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
            text="© EvKyz, 2025",
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