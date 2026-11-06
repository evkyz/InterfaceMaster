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
        self.shortcut_text_choice = tk.IntVar(value=0)

        self.create_widgets()
        self.show()

    def create_widgets(self):
        ttk.Label(self.content_frame, text="Настройка ярлыков Windows",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)

        arrows_frame = ttk.LabelFrame(self.content_frame,
                                      text="Настройка стрелок ярлыков", padding=15)
        arrows_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        ttk.Label(arrows_frame,
                  text="Выберите тип отображения стрелок на ярлыках Windows",
                  font=('Arial', 10), foreground='#7f8c8d').pack(pady=(0, 10))

        ttk.Radiobutton(arrows_frame, text="Большие стрелки (стандартные)",
                        variable=self.arrow_choice, value=1).pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(arrows_frame, text="Маленькие стрелки",
                        variable=self.arrow_choice, value=2).pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(arrows_frame, text="Убрать стрелки",
                        variable=self.arrow_choice, value=3).pack(anchor=tk.W, pady=5)

        shortcut_frame = ttk.LabelFrame(self.content_frame,
                                        text="Настройка текста ярлыков", padding=15)
        shortcut_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        ttk.Label(shortcut_frame,
                  text="Настройка отображения текста при создании ярлыков",
                  font=('Arial', 10), foreground='#7f8c8d').pack(pady=(0, 10))

        ttk.Radiobutton(shortcut_frame, text="Стандартный текст ' — ярлык'",
                        variable=self.shortcut_text_choice, value=1).pack(anchor=tk.W, pady=5)
        ttk.Radiobutton(shortcut_frame, text="Убрать текст ' — ярлык'",
                        variable=self.shortcut_text_choice, value=2).pack(anchor=tk.W, pady=5)

        button_frame = ttk.Frame(self.content_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(button_frame, text="Применить настройки",
                   command=self.apply_changes, width=25).pack(pady=(0, 10))
        ttk.Button(button_frame, text="Перезапустить Проводник",
                   command=self.restart_explorer, width=25).pack(pady=5)

        self.status_label = ttk.Label(self.content_frame, text="", font=('Arial', 9),
                                      foreground='#27ae60', wraplength=400,
                                      justify=tk.CENTER)
        self.status_label.pack(pady=5)

    def _delete_registry_key_if_exists(self, key_path, value_name):
        try:
            check_result = subprocess.run(
                ['reg', 'query', key_path, '/v', value_name],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if check_result.returncode == 0:
                subprocess.run(
                    ['reg', 'delete', key_path, '/v', value_name, '/f'],
                    capture_output=True, text=True, encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW, check=True)
                return True
        except Exception:
            pass
        return False

    def _set_registry_value(self, key_path, value_name, value_data, value_type='REG_SZ'):
        if value_data == "" and value_type == 'REG_SZ':
            cmd = ['reg', 'add', key_path, '/v', value_name, '/t', value_type, '/f']
        else:
            cmd = ['reg', 'add', key_path, '/v', value_name, '/d', value_data,
                   '/t', value_type, '/f']
        subprocess.run(cmd, capture_output=True, text=True, encoding='cp866',
                       creationflags=subprocess.CREATE_NO_WINDOW, check=True)

    def _clear_icon_cache(self):
        try:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True,
                           text=True, encoding='cp866',
                           creationflags=subprocess.CREATE_NO_WINDOW)
            return True
        except Exception:
            return False

    def apply_changes(self):
        arrow_choice = self.arrow_choice.get()
        text_choice = self.shortcut_text_choice.get()

        try:
            if arrow_choice == 1:
                self._delete_registry_key_if_exists(
                    'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
                    '29')
                self.show_status("✓ Установлены стандартные большие стрелки", False)
            elif arrow_choice == 2:
                self._set_registry_value(
                    'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
                    '29', 'C:\\Windows\\System32\\shell32.dll,-30')
                self.show_status("✓ Установлены маленькие стрелки", False)
            elif arrow_choice == 3:
                self._set_registry_value(
                    'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
                    '29', '')
                self.show_status("✓ Стрелки ярлыков удалены", False)

            if text_choice == 1:
                self._set_registry_value(
                    'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\NamingTemplates',
                    'ShortcutNameTemplate', '\\"%s — ярлык.lnk\\"')
                self.show_status("✓ Текст ' — ярлык' восстановлен", False)
            elif text_choice == 2:
                self._set_registry_value(
                    'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\NamingTemplates',
                    'ShortcutNameTemplate', '\\"%s.lnk\\"')
                self.show_status("✓ Текст ' — ярлык' удален", False)

            self._clear_icon_cache()

            if arrow_choice > 0 or text_choice > 0:
                self.show_status("✓ Настройки успешно применены")
        except subprocess.CalledProcessError:
            messagebox.showerror("Ошибка", "Не удалось изменить настройки реестра")
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Произошла непредвиденная ошибка: {str(e)}")

    def restart_explorer(self):
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                           capture_output=True, text=True, encoding='cp866',
                           creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.Popen('explorer.exe', shell=True,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            self.show_status("✓ Проводник успешно перезапущен")
        except Exception:
            messagebox.showerror("Ошибка", "Не удалось перезапустить Проводник")

    def show_status(self, message, clear_previous=True):
        if clear_previous:
            self.status_label.config(text=message, foreground='#27ae60')
        else:
            current = self.status_label.cget("text")
            self.status_label.config(
                text=f"{current}\n{message}" if current else message,
                foreground='#27ae60')
        self.frame.after(5000, lambda: self.status_label.config(text=""))

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()