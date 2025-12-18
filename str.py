import tkinter as tk
from tkinter import ttk, messagebox
import subprocess


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)

        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self.arrow_choice = tk.IntVar(value=0)

        self.create_widgets()
        self.show()

    def create_widgets(self):
        """Создание виджетов для настройки стрелок ярлыков"""
        title_label = ttk.Label(
            self.content_frame,
            text="Настройка стрелок ярлыков",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        desc_label = ttk.Label(
            self.content_frame,
            text="Выберите тип отображения стрелок на ярлыках Windows",
            font=('Arial', 10),
            foreground='#7f8c8d'
        )
        desc_label.pack(pady=(0, 20))

        settings_frame = ttk.LabelFrame(self.content_frame, text="Выбор типа стрелок", padding=15)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.big_arrow_check = ttk.Radiobutton(
            settings_frame,
            text="Большие стрелки (стандартные)",
            variable=self.arrow_choice,
            value=1
        )
        self.big_arrow_check.pack(anchor=tk.W, pady=5)

        self.small_arrow_check = ttk.Radiobutton(
            settings_frame,
            text="Маленькие стрелки",
            variable=self.arrow_choice,
            value=2
        )
        self.small_arrow_check.pack(anchor=tk.W, pady=5)

        self.no_arrow_check = ttk.Radiobutton(
            settings_frame,
            text="Убрать стрелки полностью",
            variable=self.arrow_choice,
            value=3
        )
        self.no_arrow_check.pack(anchor=tk.W, pady=5)

        button_frame = ttk.Frame(settings_frame)
        button_frame.pack(fill=tk.X, pady=15)

        apply_btn = ttk.Button(
            button_frame,
            text="Применить настройки",
            command=self.apply_changes,
            width=25
        )
        apply_btn.pack(pady=(0, 10))

        restart_btn = ttk.Button(
            button_frame,
            text="Перезапустить Проводник",
            command=self.restart_explorer,
            width=25
        )
        restart_btn.pack(pady=5)

        self.info_label = ttk.Label(
            settings_frame,
            text="После применения настроек нажмите Перезапустить Проводник",
            font=('Arial', 9),
            foreground='#3498db',
            wraplength=400,
            justify=tk.CENTER
        )
        self.info_label.pack(pady=5)

        self.status_label = ttk.Label(
            settings_frame,
            text="",
            font=('Arial', 9),
            foreground='#27ae60',
            wraplength=400,
            justify=tk.CENTER
        )
        self.status_label.pack(pady=5)

    def _delete_registry_key_if_exists(self):
        """Удалить ключ реестра, если он существует."""
        check_result = subprocess.run(
            ['reg', 'query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons', '/v', '29'],
            capture_output=True,
            text=True,
            encoding='cp866'
        )

        if check_result.returncode == 0:
            subprocess.run(
                ['reg', 'delete', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
                 '/v', '29', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                check=True
            )

    def _set_registry_value(self, value):
        """Установить значение в реестре."""
        subprocess.run(
            ['reg', 'add', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
             '/v', '29', '/d', value, '/t', 'REG_SZ', '/f'],
            capture_output=True,
            text=True,
            encoding='cp866',
            check=True
        )

    def apply_changes(self):
        """Применить выбранные изменения."""
        choice = self.arrow_choice.get()
        if choice == 0:
            messagebox.showwarning("Внимание", "Выберите один из вариантов настроек.")
            return

        try:
            if choice == 1:
                self._delete_registry_key_if_exists()
                self.show_status("✓ Установлены стандартные большие стрелки")

            elif choice == 2:
                self._set_registry_value('C:\\Windows\\System32\\shell32.dll,-30')
                self.show_status("✓ Установлены маленькие стрелки")

            elif choice == 3:
                self._set_registry_value('C:\\Windows\\System32\\shell32.dll,-50')
                self.show_status("✓ Стрелки ярлыков удалены")

        except subprocess.CalledProcessError:
            messagebox.showerror("Ошибка", "Не удалось изменить реестр")
        except Exception:
            messagebox.showerror("Ошибка", "Произошла непредвиденная ошибка")

    def restart_explorer(self):
        """Перезагрузить Проводник Windows."""
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                           capture_output=True,
                           text=True,
                           encoding='cp866')

            subprocess.Popen('explorer.exe', shell=True)

            self.show_status("✓ Проводник успешно перезагружен")

        except Exception:
            messagebox.showerror("Ошибка", "Не удалось перезагрузить Проводник")

    def show_status(self, message, is_success=True):
        """Показать статус операции"""
        color = '#27ae60' if is_success else '#e74c3c'
        self.status_label.config(text=message, foreground=color)

        self.frame.after(5000, lambda: self.status_label.config(text=""))

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()