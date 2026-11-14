import tkinter as tk
from tkinter import ttk

import logger
from core import settings_io, settings_apply, settings_reset
from core.ui_utils import add_tooltip

MOD = "about"


class AboutModule:
    def __init__(self, parent, version, dll_info=None):
        self.parent = parent
        self.version = version
        self.dll_info = dll_info or self._load_dll_info()
        self.frame = ttk.Frame(parent)
        self._create_widgets()
        self.show()

    @staticmethod
    def _load_dll_info():
        try:
            from main import DLLChecker
            return DLLChecker.get_dll_info()
        except ImportError:
            return {
                'expected_sha1': 'A522AC997631927EA8CA441F1916E9375AD19805',
                'actual_hash': None,
                'dll_path': None,
                'hash_match': False,
                'dll_found': False,
            }

    def _dll_status(self):
        if not self.dll_info['dll_found']:
            return "Не найден", '#e74c3c'
        if self.dll_info['hash_match']:
            return "Совпадает", '#27ae60'
        return "Не совпадает", '#e67e22'

    def _create_widgets(self):
        container = ttk.Frame(self.frame)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        ttk.Label(container, text="О программе",
                  font=('Arial', 18, 'bold'), foreground='#2c3e50').pack(pady=1)
        ttk.Label(container, text=f"Interface Master версия {self.version}",
                  font=('Arial', 10), foreground='#3498db').pack(pady=(0, 15))

        info = ttk.Frame(container)
        info.pack(fill=tk.BOTH, expand=True)

        self._section(info, "Описание",
                      "Программа для настройки интерфейса Windows\n"
                      "Помощник в кастомизации внешнего вида системы",
                      center=True)

        self._section(info, "Что можно настроить",
                      "• Иконки рабочего стола          • Иконки дисков\n"
                      "• Стрелки ярлыков                    • Контекстное меню\n"
                      "• Панель задач                           • Проводник\n"
                      "                            • Модификации",
                      center=False)

        self._dll_section(info)

        self._section(info, "Разработка",
                      "Идея: EvKyz\nРеализация: DeepSeek AI\n"
                      "Сборка: 13 ноября 2026\nЛицензия: MIT License",
                      font=('Arial', 8), fg='#7f8c8d', center=True)

        self._buttons(container)

    def _section(self, parent, title, text, font=('Arial', 9), fg=None, center=True):
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.pack(fill=tk.X, pady=(0, 10))
        lbl = ttk.Label(frame, text=text, font=font,
                        justify=tk.CENTER if center else tk.LEFT)
        if fg:
            lbl.config(foreground=fg)
        lbl.pack()

    def _dll_section(self, parent):
        frame = ttk.LabelFrame(parent, text="Файл imaster.dll", padding=10)
        frame.pack(fill=tk.X, pady=(0, 15))
        status_text, color = self._dll_status()
        tk.Label(frame, text=f"Статус: {status_text}",
                 font=('Consolas', 9, 'bold'), fg=color, anchor='w').pack(
            fill=tk.X, pady=(0, 5))
        if self.dll_info['dll_found'] and self.dll_info['dll_path']:
            tk.Label(frame, text=f"Путь: {self.dll_info['dll_path']}",
                     font=('Consolas', 8), anchor='w').pack(fill=tk.X, pady=(0, 3))
            if self.dll_info['actual_hash']:
                tk.Label(frame, text=f"Хеш SHA1: {self.dll_info['actual_hash']}",
                         font=('Consolas', 8), anchor='w').pack(fill=tk.X)
        else:
            tk.Label(frame, text="Файл imaster.dll не найден",
                     font=('Consolas', 8), anchor='w').pack(fill=tk.X)

    def _buttons(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, pady=(10, 5))
        left = ttk.Frame(bar)
        left.pack(side=tk.LEFT)

        save_btn = ttk.Button(left, text="Сохранить настройки",
                              command=self._on_save, width=20)
        save_btn.pack(side=tk.LEFT, padx=(0, 5))
        add_tooltip(save_btn, "Сохранить все текущие настройки в INI файл")

        load_btn = ttk.Button(left, text="Загрузить настройки",
                              command=self._on_load, width=20)
        load_btn.pack(side=tk.LEFT, padx=(0, 5))
        add_tooltip(load_btn, "Загрузить настройки из INI файла и применить их")

        reset_btn = ttk.Button(bar, text="Сбросить все настройки",
                               command=self._on_reset, width=25)
        reset_btn.pack(side=tk.RIGHT)
        add_tooltip(reset_btn,
                    "Сбросить все настройки программы к стандартным значениям Windows")

    # ---- прокси на core ----

    def _on_save(self):
        settings_io.save_settings_to_ini(self.version, parent=self.parent)

    def _on_load(self):
        config, path = settings_io.load_settings_from_ini(parent=self.parent)
        if config is None:
            return
        settings_apply.apply_settings(config, self.parent)
        from core import win_utils
        win_utils.restart_explorer("about")

    def _on_reset(self):
        settings_reset.reset_all_settings(self.parent)

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.frame.pack_forget()