import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW

SHELL_ICONS = r'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons'
NAMING_TEMPLATES = r'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\NamingTemplates'


def _reg(*args):
    return subprocess.run(['reg', *args], capture_output=True, text=True,
                          encoding='cp866', creationflags=CREATE_NO_WINDOW)


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        self.arrow_choice = tk.IntVar(value=0)
        self.shortcut_text_choice = tk.IntVar(value=0)
        self._create_widgets()
        self.show()

    def _create_widgets(self):
        ttk.Label(self.content_frame, text="Настройка ярлыков Windows",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)

        arrows = ttk.LabelFrame(self.content_frame,
                                text="Настройка стрелок ярлыков", padding=15)
        arrows.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        ttk.Label(arrows,
                  text="Выберите тип отображения стрелок на ярлыках Windows",
                  font=('Arial', 10), foreground='#7f8c8d').pack(pady=(0, 10))
        for text, val in (("Большие стрелки (стандартные)", 1),
                          ("Маленькие стрелки", 2),
                          ("Убрать стрелки", 3)):
            ttk.Radiobutton(arrows, text=text,
                            variable=self.arrow_choice, value=val).pack(anchor=tk.W, pady=5)

        shortcuts = ttk.LabelFrame(self.content_frame,
                                   text="Настройка текста ярлыков", padding=15)
        shortcuts.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        ttk.Label(shortcuts,
                  text="Настройка отображения текста при создании ярлыков",
                  font=('Arial', 10), foreground='#7f8c8d').pack(pady=(0, 10))
        for text, val in (("Стандартный текст ' — ярлык'", 1),
                          ("Убрать текст ' — ярлык'", 2)):
            ttk.Radiobutton(shortcuts, text=text,
                            variable=self.shortcut_text_choice, value=val).pack(
                anchor=tk.W, pady=5)

        btns = ttk.Frame(self.content_frame)
        btns.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(btns, text="Применить настройки", command=self.apply_changes,
                   width=25).pack(pady=(0, 10))
        ttk.Button(btns, text="Перезапустить Проводник",
                   command=self._restart_explorer, width=25).pack(pady=5)

        self.status_label = ttk.Label(self.content_frame, text="",
                                      font=('Arial', 9), foreground='#27ae60',
                                      wraplength=400, justify=tk.CENTER)
        self.status_label.pack(pady=5)

    @staticmethod
    def _delete_value_if_exists(key_path, value_name):
        r = _reg('query', key_path, '/v', value_name)
        if r.returncode == 0:
            _reg('delete', key_path, '/v', value_name, '/f')

    @staticmethod
    def _set_value(key_path, value_name, value_data, value_type='REG_SZ'):
        if value_data == "" and value_type == 'REG_SZ':
            _reg('add', key_path, '/v', value_name, '/t', value_type, '/f')
        else:
            _reg('add', key_path, '/v', value_name, '/d', value_data,
                 '/t', value_type, '/f')

    @staticmethod
    def _clear_icon_cache():
        try:
            subprocess.run(['ie4uinit.exe', '-show'], capture_output=True,
                           creationflags=CREATE_NO_WINDOW)
        except Exception:
            pass

    def apply_changes(self):
        try:
            arrow = self.arrow_choice.get()
            text = self.shortcut_text_choice.get()

            if arrow == 1:
                self._delete_value_if_exists(SHELL_ICONS, '29')
                self._show_status("✓ Установлены стандартные большие стрелки", False)
            elif arrow == 2:
                self._set_value(SHELL_ICONS, '29',
                                'C:\\Windows\\System32\\shell32.dll,-30')
                self._show_status("✓ Установлены маленькие стрелки", False)
            elif arrow == 3:
                self._set_value(SHELL_ICONS, '29', '')
                self._show_status("✓ Стрелки ярлыков удалены", False)

            if text == 1:
                self._set_value(NAMING_TEMPLATES, 'ShortcutNameTemplate',
                                '\\"%s — ярлык.lnk\\"')
                self._show_status("✓ Текст ' — ярлык' восстановлен", False)
            elif text == 2:
                self._set_value(NAMING_TEMPLATES, 'ShortcutNameTemplate',
                                '\\"%s.lnk\\"')
                self._show_status("✓ Текст ' — ярлык' удален", False)

            self._clear_icon_cache()
            if arrow > 0 or text > 0:
                self._show_status("✓ Настройки успешно применены")
        except subprocess.CalledProcessError:
            messagebox.showerror("Ошибка", "Не удалось изменить настройки реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка: {e}")

    def _restart_explorer(self):
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                           capture_output=True, creationflags=CREATE_NO_WINDOW)
            subprocess.Popen('explorer.exe', shell=True,
                             creationflags=CREATE_NO_WINDOW)
            self._show_status("✓ Проводник успешно перезапущен")
        except Exception:
            messagebox.showerror("Ошибка", "Не удалось перезапустить Проводник")

    def _show_status(self, message, clear_previous=True):
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