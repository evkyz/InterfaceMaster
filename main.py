import hashlib
import os
import platform
import shutil
import sys
import tkinter as tk
import winreg
from tkinter import ttk, messagebox

from PIL import Image, ImageTk

VALID_HASHES = [
    'F62A4326A8300D9B828744E24314B98433510195',
    'A522AC997631927EA8CA441F1916E9375AD19805',
]
REG_ROOT = r"Software\InterfaceMaster"
REG_DLL_PATH = "DLLPath"


def calculate_sha1(filepath):
    try:
        h = hashlib.sha1()
        with open(filepath, 'rb') as f:
            for block in iter(lambda: f.read(4096), b""):
                h.update(block)
        return h.hexdigest().upper()
    except Exception:
        return None


def check_windows_version():
    system = platform.system()
    version = platform.version()
    if system != "Windows":
        return False, f"Не Windows система: {system}"
    try:
        parts = version.split('.')
        major = int(parts[0])
        minor = int(parts[1])
        build = int(parts[2]) if len(parts) > 2 else 0
        if major == 10 and minor == 0 and build >= 22000:
            return False, f"Windows 11 (сборка {build})"
        if major == 10 and minor == 0:
            return True, "Windows 10"
        return False, f"Windows {major}.{minor} (сборка {build})"
    except (ValueError, IndexError):
        return False, f"Windows (неизвестная версия: {version})"


class DLLChecker:
    dll_info = None
    copy_offered = False

    @staticmethod
    def save_dll_path_to_registry(dll_path):
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_ROOT) as k:
                winreg.SetValueEx(k, REG_DLL_PATH, 0, winreg.REG_SZ, dll_path)
            return True
        except Exception:
            return False

    @staticmethod
    def get_dll_path_from_registry():
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_ROOT) as k:
                path, _ = winreg.QueryValueEx(k, REG_DLL_PATH)
                return path if os.path.exists(path) else None
        except (FileNotFoundError, OSError):
            return None

    @staticmethod
    def initialize():
        if DLLChecker.dll_info is None:
            base = sys._MEIPASS if getattr(sys, 'frozen', False) \
                else os.path.dirname(os.path.abspath(__file__))
            DLLChecker.dll_info = DLLCheckerInternal(base).check_imaster_dll()

    @staticmethod
    def get_dll_info():
        if DLLChecker.dll_info is None:
            DLLChecker.initialize()
        return DLLChecker.dll_info.copy()

    @staticmethod
    def copy_dll_to_system32(dll_path):
        try:
            if not dll_path or not os.path.exists(dll_path):
                return False, "Исходный файл DLL не найден"
            target = r"C:\Windows\System32\imaster.dll"
            shutil.copy2(dll_path, target)
            if os.path.exists(target):
                DLLChecker.save_dll_path_to_registry(target)
                DLLChecker.initialize()
                return True, "Файл успешно скопирован в System32"
            return False, "Не удалось скопировать файл"
        except PermissionError:
            return False, "Отказано в доступе. Запустите от имени администратора."
        except Exception as e:
            return False, f"Ошибка при копировании: {e}"

    @staticmethod
    def get_dll_status():
        info = DLLChecker.get_dll_info()
        sys32 = r"C:\Windows\System32\imaster.dll"
        sys32_exists = os.path.exists(sys32)
        sys32_match = sys32_exists and calculate_sha1(sys32) in VALID_HASHES

        saved = DLLChecker.get_dll_path_from_registry()
        saved_valid = bool(saved) and calculate_sha1(saved) in VALID_HASHES

        if saved_valid:
            recommended, source = saved, "saved_registry"
        elif sys32_exists and sys32_match:
            recommended, source = sys32, "system32"
        elif info['dll_found'] and info['hash_match']:
            recommended, source = info['dll_path'], "found"
        else:
            recommended, source = None, "none"

        return {
            **info,
            'system32_path': sys32,
            'system32_exists': sys32_exists,
            'system32_hash_match': sys32_match,
            'recommended_path': recommended,
            'should_offer_copy': (info['dll_found'] and info['hash_match']
                                  and (not sys32_exists or not sys32_match)),
            'is_in_system32': sys32_exists and sys32_match,
            'saved_path': saved,
            'source': source,
            'saved_path_valid': saved_valid,
        }

    @staticmethod
    def offer_copy_to_system32(parent, module_name):
        status = DLLChecker.get_dll_status()
        if not status['should_offer_copy'] or DLLChecker.copy_offered:
            return status['recommended_path']
        DLLChecker.copy_offered = True

        msg = ("Файл imaster.dll в System32 повреждён!\n\n"
               "Хотите заменить файл в C:\\Windows\\System32?"
               if status['system32_exists'] and not status['system32_hash_match']
               else "Хотите скопировать файл в C:\\Windows\\System32?")

        if messagebox.askyesno("Копирование DLL", msg):
            ok, msg2 = DLLChecker.copy_dll_to_system32(status['dll_path'])
            if ok:
                if msg2:
                    messagebox.showinfo("Успех", msg2)
                DLLChecker.initialize()
                DLLChecker.save_dll_path_to_registry(status['system32_path'])
                return status['system32_path']
            if msg2:
                messagebox.showwarning("Предупреждение",
                                       f"Не удалось скопировать файл:\n{msg2}\n\n"
                                       f"Будет использовано текущее расположение")
            DLLChecker.save_dll_path_to_registry(status['dll_path'])
            return status['dll_path']
        DLLChecker.save_dll_path_to_registry(status['dll_path'])
        return status['dll_path']

    @staticmethod
    def get_dll_path_for_module(parent, module_name):
        status = DLLChecker.get_dll_status()
        if not status['dll_found']:
            if parent and module_name in ("desktop", "disk"):
                messagebox.showwarning("Внимание", "Файл imaster.dll не найден!")
            return None
        if not status['hash_match']:
            if parent and module_name in ("desktop", "disk"):
                messagebox.showwarning("Внимание", "Файл imaster.dll поврежден!")
            return None
        if module_name in ("desktop", "disk"):
            return DLLChecker.offer_copy_to_system32(parent, module_name)
        return status['recommended_path']

    @staticmethod
    def get_dll_path_without_prompt():
        status = DLLChecker.get_dll_status()
        if not status['dll_found'] or not status['hash_match']:
            return None
        return status['recommended_path']


class DLLCheckerInternal:
    VALID_HASHES = VALID_HASHES

    def __init__(self, base_path):
        self.base_path = base_path

    def search_dll_in_filesystem(self):
        locations = [
            (r"C:\Windows\System32\imaster.dll", "System32"),
            (os.path.join(self.base_path, "imaster.dll"), "Папка exe"),
            (os.path.join(os.getcwd(), "imaster.dll"), "Рабочая папка"),
            (os.path.join(os.path.dirname(self.base_path), "imaster.dll"),
             "Родительская от exe"),
            (os.path.join(os.path.dirname(os.getcwd()), "imaster.dll"),
             "Родительская от рабочей"),
            (r"C:\Windows\SysWOW64\imaster.dll", "SysWOW64"),
            (r"C:\Windows\imaster.dll", "Windows"),
            (r"C:\imaster.dll", "Корень C"),
        ]
        for path, _ in locations:
            if os.path.exists(path):
                h = calculate_sha1(path)
                if h and h in self.VALID_HASHES:
                    return path
        return None

    def verify_dll_hash(self, path):
        actual = calculate_sha1(path)
        return {
            'valid_hashes': self.VALID_HASHES,
            'actual_hash': actual,
            'dll_path': path,
            'hash_match': actual in self.VALID_HASHES if actual else False,
            'dll_found': True,
        }

    def check_imaster_dll(self):
        saved = DLLChecker.get_dll_path_from_registry()
        if saved and os.path.exists(saved):
            r = self.verify_dll_hash(saved)
            if r['hash_match']:
                return r
        found = self.search_dll_in_filesystem()
        if found:
            return self.verify_dll_hash(found)
        return {
            'valid_hashes': self.VALID_HASHES,
            'actual_hash': None,
            'dll_path': None,
            'hash_match': False,
            'dll_found': False,
        }


class InterfaceMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface Master v2.1 Beta")
        self.root.geometry("520x710")
        self.root.resizable(False, False)

        self._windows_version_check()

        self.base_path = sys._MEIPASS if getattr(sys, 'frozen', False) \
            else os.path.dirname(os.path.abspath(__file__))
        if self.base_path not in sys.path:
            sys.path.insert(0, self.base_path)

        icon_path = os.path.join(self.base_path, "icon.ico")
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        DLLChecker.initialize()

        self.version = "2.1.3.185 Beta"
        self.logo_photo = None
        self.modules_loaded = self._load_modules()
        self.current_module_instance = None

        self._center_window()
        self._setup_styles()
        self._load_logo()
        self._create_interface()

    def _windows_version_check(self):
        is_w10, _ = check_windows_version()
        if not is_w10:
            self.root.after(100, lambda: messagebox.showwarning(
                "Проверка системы",
                "Внимание! Корректная работа программы\n"
                "на данной системе не гарантируется!"))

    def _load_modules(self):
        modules = {}
        names = ('about', 'taskbar', 'disk', 'mod', 'str',
                 'desktop', 'contextmenu', 'explorer')
        for name in names:
            try:
                if getattr(sys, 'frozen', False):
                    import importlib.util
                    path = os.path.join(self.base_path, f"{name}.py")
                    if os.path.exists(path):
                        spec = importlib.util.spec_from_file_location(f"imod_{name}", path)
                        mod = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(mod)
                        modules[name] = mod
                else:
                    modules[name] = __import__(name)
            except Exception:
                continue
        return modules

    def _load_logo(self):
        path = os.path.join(self.base_path, "logo.png")
        try:
            if os.path.exists(path):
                img = Image.open(path).resize((40, 40), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(img)
        except Exception:
            self.logo_photo = None

    def _center_window(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'{w}x{h}+{x}+{y}')

    def _setup_styles(self):
        s = ttk.Style()
        s.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        s.configure('Main.TButton', font=('Arial', 9), padding=10)
        s.configure('Back.TButton', font=('Arial', 8), padding=4, foreground='#34495e')
        s.configure('Help.TButton', font=('Arial', 10, 'bold'), padding=2, width=3)

    def _create_interface(self):
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        self.header_frame = ttk.Frame(self.main_container)
        self.header_frame.pack(fill=tk.X, pady=(10, 5))
        self.content_frame = ttk.Frame(self.main_container)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.footer_frame = ttk.Frame(self.main_container)
        self.footer_frame.pack(fill=tk.X, pady=(5, 10))

        self._create_header()
        self._create_menu()
        self._create_footer()
        self.module_frame = ttk.Frame(self.content_frame)

    def _create_header(self):
        content = ttk.Frame(self.header_frame)
        content.pack(expand=True, fill=tk.X, padx=20)
        if self.logo_photo:
            ttk.Label(content, image=self.logo_photo).grid(
                row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(content, text="Interface Master",
                  style='Title.TLabel').grid(row=0, column=1)
        ttk.Button(content, text="?", style='Help.TButton',
                   command=self._open_about).grid(row=0, column=2, sticky="e", padx=(10, 0))
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=0)
        content.grid_columnconfigure(2, weight=1)

    def _create_menu(self):
        self.menu_frame = ttk.Frame(self.content_frame)
        self.menu_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        buttons = [
            ("Настройка иконок Рабочего стола", self._open_desktop),
            ("Настройка иконок дисков", self._open_disk),
            ("Настройка стрелок ярлыков", self._open_arrows),
            ("Настройка контекстного меню", self._open_context_menu),
            ("Настройка Панели задач", self._open_taskbar),
            ("Проводник", self._open_explorer),
            ("Модификации", self._open_modifications),
        ]
        for text, cmd in buttons:
            ttk.Button(self.menu_frame, text=text, style='Main.TButton',
                       command=cmd).pack(fill=tk.X, pady=3, ipady=6)

    def _create_footer(self):
        left = ttk.Frame(self.footer_frame)
        left.pack(side=tk.LEFT, padx=10)
        self.back_btn = ttk.Button(left, text="Назад", style='Back.TButton',
                                   command=self._back_to_menu, width=15)
        self.back_btn.grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.back_btn.grid_remove()
        ttk.Label(left, text=f"v{self.version}", font=('Arial', 7),
                  foreground='#95a5a6').grid(row=1, column=0, sticky="w")

        right = ttk.Frame(self.footer_frame)
        right.pack(side=tk.RIGHT, padx=10)
        ttk.Label(right, text="© EvKyz, 2026", font=('Arial', 7),
                  foreground='#95a5a6').pack(side=tk.RIGHT)

    def _open_about(self):
        self._open_module("about")

    def _open_module(self, module_name):
        if module_name not in self.modules_loaded:
            messagebox.showerror("Ошибка", f"Модуль {module_name} не загружен")
            return
        self.menu_frame.pack_forget()
        self.module_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.back_btn.grid()
        for w in self.module_frame.winfo_children():
            w.destroy()
        try:
            if module_name == 'about':
                self.current_module_instance = self.modules_loaded[module_name].AboutModule(
                    parent=self.module_frame, version=self.version,
                    dll_info=DLLChecker.get_dll_status())
            elif module_name == 'explorer':
                self.current_module_instance = self.modules_loaded[module_name].ExplorerModule(
                    parent=self.module_frame)
            else:
                cls = getattr(self.modules_loaded[module_name], 'Module', None)
                if not cls:
                    ttk.Label(self.module_frame,
                              text=f"Модуль {module_name} не имеет правильной структуры",
                              foreground='red').pack(expand=True)
                    return
                if module_name in ("desktop", "disk"):
                    self.current_module_instance = cls(
                        parent=self.module_frame, module_name=module_name)
                else:
                    self.current_module_instance = cls(parent=self.module_frame)
        except Exception:
            ttk.Label(self.module_frame,
                      text=f"Ошибка загрузки модуля {module_name}",
                      foreground='red').pack(expand=True)

    def _back_to_menu(self):
        self.module_frame.pack_forget()
        for w in self.module_frame.winfo_children():
            w.destroy()
        self.current_module_instance = None
        self.back_btn.grid_remove()
        self.menu_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    def _open_desktop(self): self._open_module("desktop")
    def _open_disk(self): self._open_module("disk")
    def _open_arrows(self): self._open_module("str")
    def _open_context_menu(self): self._open_module("contextmenu")
    def _open_taskbar(self): self._open_module("taskbar")
    def _open_modifications(self): self._open_module("mod")
    def _open_explorer(self): self._open_module("explorer")


def main():
    root = tk.Tk()
    InterfaceMaster(root)
    root.mainloop()


if __name__ == "__main__":
    main()