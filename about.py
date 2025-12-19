import tkinter as tk
from tkinter import ttk


class AboutModule:
    def __init__(self, parent, version, dll_info=None):
        self.parent = parent
        self.version = version

        self.dll_info = dll_info
        if self.dll_info is None:
            try:
                from main import DLLChecker
                self.dll_info = DLLChecker.get_dll_info()
            except ImportError:
                self.dll_info = {
                    'expected_sha1': 'F62A4326A8300D9B828744E24314B98433510195',
                    'actual_hash': None,
                    'dll_path': None,
                    'hash_match': False,
                    'dll_found': False
                }

        self.frame = ttk.Frame(parent)
        self.create_widgets()
        self.show()

    def get_dll_status_text(self):
        if not self.dll_info['dll_found']:
            return "Не найден"

        if self.dll_info['hash_match']:
            return "Совпадает"
        else:
            return "Не совпадает"

    def get_dll_status_color(self):
        if not self.dll_info['dll_found']:
            return '#e74c3c'

        if self.dll_info['hash_match']:
            return '#27ae60'
        else:
            return '#e67e22'

    def create_widgets(self):
        title_label = ttk.Label(
            self.frame,
            text="О программе",
            font=('Arial', 18, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=1)

        main_container = ttk.Frame(self.frame)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        version_label = ttk.Label(
            main_container,
            text=f"Interface Master версия {self.version}",
            font=('Arial', 10),
            foreground='#3498db'
        )
        version_label.pack(pady=(0, 15))

        desc_frame = ttk.LabelFrame(
            main_container,
            text="Описание",
            padding=10
        )
        desc_frame.pack(fill=tk.X, pady=(0, 10))

        desc_text = """Программа для настройки интерфейса Windows.
Помощник в кастомизации внешнего вида системы."""

        desc_label = ttk.Label(
            desc_frame,
            text=desc_text,
            font=('Arial', 9),
            justify=tk.CENTER
        )
        desc_label.pack()

        features_frame = ttk.LabelFrame(
            main_container,
            text="Что можно настроить",
            padding=10
        )
        features_frame.pack(fill=tk.X, pady=(0, 15))

        features_text = """• Иконки рабочего стола          • Иконки дисков
• Стрелки ярлыков                    • Контекстное меню
• Панель задач                           • Проводник
                            • Модификации"""

        features_label = ttk.Label(
            features_frame,
            text=features_text,
            font=('Arial', 9),
            justify=tk.LEFT
        )
        features_label.pack()

        dll_frame = ttk.LabelFrame(
            main_container,
            text="Файл imaster.dll",
            padding=10
        )
        dll_frame.pack(fill=tk.X, pady=(0, 15))

        self.create_dll_section(dll_frame)

        dev_frame = ttk.LabelFrame(
            main_container,
            text="Разработка",
            padding=10
        )
        dev_frame.pack(fill=tk.X, pady=(0, 10))

        dev_text = f"""Идея: EvKyz
Реализация: DeepSeek AI
Сборка: 19 декабря 2025
Лицензия: MIT License"""

        dev_label = ttk.Label(
            dev_frame,
            text=dev_text,
            font=('Arial', 8),
            justify=tk.CENTER,
            foreground='#7f8c8d'
        )
        dev_label.pack()

    def create_dll_section(self, dll_frame):
        container = tk.Frame(dll_frame, bg='#f5f5f5', padx=5, pady=5)
        container.pack(fill=tk.BOTH, expand=True)

        status_color = self.get_dll_status_color()
        status_text = self.get_dll_status_text()

        status_label = tk.Label(
            container,
            text=f"Статус: {status_text}",
            font=('Consolas', 9, 'bold'),
            fg=status_color,
            bg='#f5f5f5',
            anchor='w',
            justify=tk.LEFT
        )
        status_label.pack(fill=tk.X, pady=(0, 5))

        if self.dll_info['dll_found'] and self.dll_info['dll_path']:
            path_label = tk.Label(
                container,
                text=f"Путь: {self.dll_info['dll_path']}",
                font=('Consolas', 8),
                bg='#f5f5f5',
                anchor='w',
                justify=tk.LEFT
            )
            path_label.pack(fill=tk.X, pady=(0, 3))

            if self.dll_info['actual_hash']:
                hash_label = tk.Label(
                    container,
                    text=f"Хеш SHA1: {self.dll_info['actual_hash']}",
                    font=('Consolas', 8),
                    bg='#f5f5f5',
                    anchor='w',
                    justify=tk.LEFT
                )
                hash_label.pack(fill=tk.X, pady=(0, 3))
        else:
            not_found_label = tk.Label(
                container,
                text="Файл imaster.dll не найден",
                font=('Consolas', 8),
                bg='#f5f5f5',
                anchor='w',
                justify=tk.LEFT
            )
            not_found_label.pack(fill=tk.X, pady=(0, 3))

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()