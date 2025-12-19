import os
import string
import subprocess
import tkinter as tk
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
        self.module_name = module_name or "disk"

        self.dll_path = self.get_dll_path()

        self.icon_packages = {
            "Windows 95": {
                "Дискета": 101,
                "Дисковод": 102,
                "Жесткий диск": 103,
                "Флешка": 105,
            },
            "Windows 98": {
                "Дискета": 101,
                "Дисковод": 102,
                "Жесткий диск": 103,
                "Флешка": 105,
            },
            "Windows 2000": {
                "Дискета": 101,
                "Дисковод": 102,
                "Жесткий диск": 103,
                "Флешка": 105,
            },
            "Windows XP": {
                "Дискета": 401,
                "Дисковод": 402,
                "Жесткий диск": 403,
                "Жесткий диск (sys)": 404,
                "Флешка": 405,
            },
            "Windows Vista": {
                "Дискета": 501,
                "Дисковод": 502,
                "Жесткий диск": 503,
                "Жесткий диск (sys)": 504,
                "Флешка": 505,
            },
            "Windows 7": {
                "Дискета": 501,
                "Дисковод": 502,
                "Жесткий диск": 503,
                "Жесткий диск (sys)": 504,
                "Флешка": 505,
            },
            "Windows 10": {
                "Дискета": 701,
                "Дисковод": 702,
                "Жесткий диск": 703,
                "Жесткий диск (sys)": 704,
                "Флешка": 705,
            },
            "Windows 11": {
                "Дискета": 801,
                "Дисковод": 802,
                "Жесткий диск": 803,
                "Жесткий диск (sys)": 804,
                "Флешка": 805,
            },
        }

        self.message_timer = None

        self.create_widgets()
        self.show()

    def get_dll_path(self):
        if not HAS_DLL_CHECKER:
            # Не показываем сообщение - в main уже есть проверка
            return None

        dll_path = DLLChecker.get_dll_path_for_module(self.parent, self.module_name)
        return dll_path

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

    def get_available_disks(self):
        disks = []
        for drive_letter in string.ascii_uppercase:
            drive_path = f"{drive_letter}:\\"
            if os.path.exists(drive_path):
                disks.append(drive_path)
        return disks

    def create_widgets(self):
        title_label = ttk.Label(
            self.frame,
            text="Настройка иконок дисков",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        settings_frame = ttk.LabelFrame(self.frame, text="Настройки иконок дисков", padding=15)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        disk_frame = ttk.Frame(settings_frame)
        disk_frame.pack(fill=tk.X, pady=8)

        self.disk_label = ttk.Label(disk_frame, text="Выберите диск:", font=('Arial', 9), width=15)
        self.disk_label.pack(side=tk.LEFT, padx=(0, 10))

        self.disk_var = tk.StringVar()
        self.disk_combobox = ttk.Combobox(disk_frame, textvariable=self.disk_var, width=20)
        self.disk_combobox['values'] = self.get_available_disks()
        self.disk_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if self.disk_combobox['values']:
            self.disk_combobox.current(0)

        package_frame = ttk.Frame(settings_frame)
        package_frame.pack(fill=tk.X, pady=8)

        self.package_label = ttk.Label(package_frame, text="Выберите пакет:", font=('Arial', 9), width=15)
        self.package_label.pack(side=tk.LEFT, padx=(0, 10))

        self.package_var = tk.StringVar()
        self.package_combobox = ttk.Combobox(package_frame, textvariable=self.package_var, width=20)
        self.package_combobox['values'] = [
            "Windows 95", "Windows 98", "Windows 2000", "Windows XP",
            "Windows Vista", "Windows 7", "Windows 10", "Windows 11"
        ]
        self.package_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.package_combobox.current(0)
        self.package_combobox.bind("<<ComboboxSelected>>", self.update_icon_combobox)

        icon_frame = ttk.Frame(settings_frame)
        icon_frame.pack(fill=tk.X, pady=8)

        self.icon_label = ttk.Label(icon_frame, text="Выберите иконку:", font=('Arial', 9), width=15)
        self.icon_label.pack(side=tk.LEFT, padx=(0, 10))

        self.icon_var = tk.StringVar()
        self.icon_combobox = ttk.Combobox(icon_frame, textvariable=self.icon_var, width=20)
        self.icon_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.update_icon_combobox()

        apply_frame = ttk.Frame(settings_frame)
        apply_frame.pack(fill=tk.X, pady=15)

        self.apply_button = ttk.Button(
            apply_frame,
            text="Применить иконку",
            command=self.apply_icon,
            state='normal' if self.dll_path else 'disabled'
        )
        self.apply_button.pack(fill=tk.X, pady=2)

        delete_frame = ttk.Frame(settings_frame)
        delete_frame.pack(fill=tk.X, pady=5)

        self.remove_icon_button = ttk.Button(
            delete_frame,
            text="Удалить иконку",
            command=self.remove_icon
        )
        self.remove_icon_button.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)

        self.remove_all_icons_button = ttk.Button(
            delete_frame,
            text="Удалить все иконки",
            command=self.remove_all_icons
        )
        self.remove_all_icons_button.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

        self.message_label = ttk.Label(
            settings_frame,
            text="",
            font=('Arial', 9),
            wraplength=400
        )

    def update_icon_combobox(self, event=None):
        selected_package = self.package_var.get()
        if selected_package in self.icon_packages:
            icons = list(self.icon_packages[selected_package].keys())
            self.icon_combobox['values'] = icons
            if icons:
                self.icon_combobox.current(0)

    def apply_icon(self):
        # Проверка DLL уже сделана при загрузке, кнопка будет отключена если нет DLL
        if not self.dll_path:
            return

        selected_disk = self.disk_var.get()
        selected_package = self.package_var.get()
        selected_icon = self.icon_var.get()

        if selected_disk and selected_package and selected_icon:
            selected_disk = selected_disk.strip(":\\/").upper()
            icon_number = self.icon_packages[selected_package].get(selected_icon)

            if selected_disk and icon_number:
                try:
                    subprocess.run(
                        [
                            'reg', 'add',
                            f'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons\\{selected_disk}\\DefaultIcon',
                            '/ve', '/d', f'{self.dll_path},-{icon_number}',
                            '/t', 'REG_SZ', '/f'
                        ],
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        encoding='cp866'
                    )

                    self.show_message(f"Иконка для диска {selected_disk} установлена", "success")

                except subprocess.CalledProcessError:
                    # Не показываем сообщение об ошибке
                    pass

    def remove_icon(self):
        selected_disk = self.disk_var.get()
        if selected_disk:
            selected_disk = selected_disk.strip(":\\/").upper()

            if selected_disk:
                try:
                    subprocess.run(
                        [
                            'reg', 'delete',
                            f'HKLM\\SOFTWARE\\Microsoft\\Windows\CurrentVersion\\Explorer\\DriveIcons\\{selected_disk}',
                            '/f'
                        ],
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        encoding='cp866'
                    )
                    self.show_message(f"Иконка для диска {selected_disk} удалена", "success")

                except subprocess.CalledProcessError:
                    # Не показываем сообщение об ошибке
                    pass

    def remove_all_icons(self):
        try:
            subprocess.run(
                [
                    'reg', 'delete',
                    'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\Explorer\\DriveIcons',
                    '/f'
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='cp866'
            )
            self.show_message("Все иконки дисков удалены", "success")

        except subprocess.CalledProcessError:
            # Не показываем сообщение об ошибке
            pass

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()