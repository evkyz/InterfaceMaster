import os
import re
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW

FOLDERS = {
    "{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}": ["Видео", "Папка с видео файлами"],
    "{088e3905-0323-4b02-9826-5d99428e115f}": ["Загрузки", "Папка для загруженных файлов"],
    "{d3162b92-9365-467a-956b-92703aca08af}": ["Документы", "Папка для документов"],
    "{24ad3ad4-a569-4530-98e1-ab02f9417aa8}": ["Изображения", "Папка с изображениями"],
    "{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}": ["Музыка", "Папка с аудио файлами"],
    "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}": ["Рабочий стол", "Папка рабочего стола"],
    "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}": ["3D объекты", "Папка для 3D моделей"],
}

NAMESPACE_KEY = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"
IMAGE_DETAILS_VALUE = "{2cc5ca98-6485-489a-920e-b3e88a6ccce3}"
IMAGE_DETAILS_KEY = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"

PHOTO_VIEWER_EXTS = {
    '.jpg': 'PhotoViewer.FileAssoc.Jpeg',
    '.wdp': 'PhotoViewer.FileAssoc.Wdp',
    '.jfif': 'PhotoViewer.FileAssoc.JFIF',
    '.dib': 'PhotoViewer.FileAssoc.Bitmap',
    '.png': 'PhotoViewer.FileAssoc.Png',
    '.tiff': 'PhotoViewer.FileAssoc.Tiff',
    '.jxr': 'PhotoViewer.FileAssoc.Wdp',
    '.bmp': 'PhotoViewer.FileAssoc.Bitmap',
    '.jpe': 'PhotoViewer.FileAssoc.Jpeg',
    '.jpeg': 'PhotoViewer.FileAssoc.Jpeg',
    '.gif': 'PhotoViewer.FileAssoc.Gif',
    '.tif': 'PhotoViewer.FileAssoc.Tiff',
}

PHOTO_VIEWER_CLASSES = [
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap',
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF',
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg',
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png',
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp',
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Gif',
    r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Tiff',
]


def _run_reg(*args):
    return subprocess.run(['reg', *args], capture_output=True, text=True,
                          encoding='cp866', creationflags=CREATE_NO_WINDOW)


class ExplorerModule:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.checkbox_vars = {}
        self.bitlocker_enabled = False
        self.photo_viewer_enabled = False
        self.file_ext_enabled = False
        self.recent_folders_enabled = True
        self.image_details_hidden = False
        self.image_details_supported = False
        self.status_labels = {}

        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.main_container = tk.Frame(self.content_frame)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        self.main_canvas = tk.Canvas(self.main_container, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self.main_container, orient="vertical",
                                            command=self.main_canvas.yview)
        self.main_scrollable = ttk.Frame(self.main_canvas)
        self.main_scrollable.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_window = self.main_canvas.create_window(
            (0, 0), window=self.main_scrollable, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        self.main_canvas.bind(
            "<Configure>",
            lambda e: self.main_canvas.itemconfigure(self.main_window, width=e.width))

        self.main_menu_frame = ttk.Frame(self.main_scrollable)
        self.folders_frame = ttk.Frame(self.main_container)

        self._create_main_menu()
        self._create_folders_screen()
        self._show_main_menu()
        self._check_initial_states()
        self.show()

    # ---- helpers ----
    @staticmethod
    def _is_windows_10_22h2_plus():
        try:
            r = _run_reg('query', r'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion',
                         '/v', 'CurrentBuildNumber')
            if r.returncode == 0:
                for line in r.stdout.splitlines():
                    if 'CurrentBuildNumber' in line:
                        build = int(line.split('REG_SZ')[1].strip())
                        return 19045 <= build < 22000
        except Exception:
            pass
        try:
            v = sys.getwindowsversion()
            return v.major == 10 and 19045 <= v.build < 22000
        except Exception:
            pass
        try:
            import platform
            parts = platform.version().split('.')
            if len(parts) >= 3:
                build = int(parts[2])
                return 19045 <= build < 22000
        except Exception:
            pass
        return False

    def _add_tooltip(self, widget, text):
        state = {'tip': None}

        def enter(event):
            if state['tip']:
                return
            try:
                x = widget.winfo_rootx() + 20
                y = widget.winfo_rooty() + widget.winfo_height() + 5
                tip = tk.Toplevel(widget)
                tip.wm_overrideredirect(True)
                tip.wm_geometry(f"+{x}+{y}")
                tk.Label(tip, text=text, background="#ffffe0", relief="solid",
                         borderwidth=1, font=("Arial", 9)).pack()
                state['tip'] = tip
            except Exception:
                state['tip'] = None

        def leave(event):
            if state['tip']:
                try:
                    state['tip'].destroy()
                except Exception:
                    pass
                state['tip'] = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    # ---- main menu ----
    def _create_main_menu(self):
        f = self.main_menu_frame
        f.pack(fill=tk.BOTH, expand=True, padx=(20, 35), pady=10)

        ttk.Label(f, text="Настройки Проводника",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=(0, 30))

        ttk.Button(f, text="Папки в Этот компьютер",
                   command=self._show_folders_screen, width=35).pack(pady=10)

        clear_btn = ttk.Button(f, text="Очистка истории",
                               command=self.clear_typed_paths, width=35)
        clear_btn.pack(pady=(0, 10))
        self._add_tooltip(clear_btn, "Для выборочного удаления зажмите Shift")
        clear_btn.bind("<Shift-Button-1>", self._on_clear_history_shift_click)

        ttk.Separator(f, orient='horizontal').pack(fill=tk.X, pady=20)

        self._add_toggle_section(f, "Управление пунктом BitLocker в контекстном меню дисков:",
                                 "bitlocker_status", "bitlocker_btn", self.toggle_bitlocker)
        ttk.Separator(f, orient='horizontal').pack(fill=tk.X, pady=20)
        self._add_toggle_section(f, "Включение/отключение Windows Photo Viewer:",
                                 "photo_viewer_status", "photo_viewer_btn",
                                 self.toggle_photo_viewer)
        ttk.Separator(f, orient='horizontal').pack(fill=tk.X, pady=20)
        self._add_toggle_section(f, "Показ расширений файлов в Проводнике:",
                                 "file_ext_status", "file_ext_btn", self.toggle_file_ext)
        ttk.Separator(f, orient='horizontal').pack(fill=tk.X, pady=20)
        self._add_toggle_section(f, "Отображение недавних папок в Проводнике:",
                                 "recent_folders_status", "recent_folders_btn",
                                 self.toggle_recent_folders)
        ttk.Separator(f, orient='horizontal').pack(fill=tk.X, pady=20)

        # Image details
        frame = ttk.Frame(f)
        frame.pack(fill=tk.X, pady=5)
        ttk.Label(frame,
                  text="Отображение значка \nДополнительные сведения об этом изображении:",
                  font=('Arial', 9), foreground='#7f8c8d').pack(anchor=tk.W)
        self.image_details_status = ttk.Label(frame, text="Состояние: проверяется...",
                                              font=('Arial', 9), foreground='#95a5a6')
        self.image_details_status.pack(anchor=tk.W, pady=(2, 5))
        self.image_details_btn = ttk.Button(frame, text="Включить значок",
                                            command=self.toggle_image_details, width=35)
        self.image_details_btn.pack()

        self.image_details_supported = self._is_windows_10_22h2_plus()
        if not self.image_details_supported:
            self.image_details_btn.config(state=tk.DISABLED)
            self._add_tooltip(self.image_details_btn,
                              "Эта функция доступна только в Windows 10 22H2")

    def _add_toggle_section(self, parent, label, status_attr, btn_attr, command):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        ttk.Label(frame, text=label, font=('Arial', 9),
                  foreground='#7f8c8d').pack(anchor=tk.W)
        status = ttk.Label(frame, text="Состояние: проверяется...",
                           font=('Arial', 9), foreground='#95a5a6')
        status.pack(anchor=tk.W, pady=(2, 5))
        setattr(self, status_attr, status)
        btn = ttk.Button(frame, text="---", command=command, width=35)
        btn.pack()
        setattr(self, btn_attr, btn)

    # ---- folders screen ----
    def _create_folders_screen(self):
        f = self.folders_frame
        header = ttk.Frame(f)
        header.pack(fill=tk.X, pady=(0, 20))
        ttk.Button(header, text="← Назад",
                   command=self._show_main_menu, width=10).pack(side=tk.LEFT)
        ttk.Label(header, text="Папки в Этот компьютер",
                  font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(
            side=tk.LEFT, padx=20)

        ttk.Label(f, text="Выберите какие папки отображать в Этот компьютер",
                  font=('Arial', 10), foreground='#7f8c8d').pack(pady=(0, 15))

        list_frame = ttk.LabelFrame(f, text="Доступные папки", padding=15)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        canvas = tk.Canvas(list_frame, highlightthickness=0, height=220)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for guid, (name, desc) in FOLDERS.items():
            row = ttk.Frame(inner)
            row.pack(fill=tk.X, padx=5, pady=4)
            var = tk.BooleanVar()
            self.checkbox_vars[guid] = var
            ttk.Checkbutton(row, text=name, variable=var,
                            onvalue=True, offvalue=False).pack(side=tk.LEFT)
            status = ttk.Label(row, text="•", font=('Arial', 9, 'bold'), width=2)
            status.pack(side=tk.LEFT, padx=(2, 0))
            self.status_labels[guid] = status
            ttk.Label(row, text=f" - {desc}", font=('Arial', 9),
                      foreground='#95a5a6').pack(side=tk.LEFT, padx=(0, 8))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        top = ttk.Frame(f)
        top.pack(pady=15)
        ttk.Button(top, text="Выбрать все папки",
                   command=self._select_all, width=20).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(top, text="Снять все папки",
                   command=self._deselect_all, width=20).pack(side=tk.LEFT)

        bottom = ttk.Frame(f)
        bottom.pack(pady=(0, 15))
        ttk.Button(bottom, text="Применить изменения",
                   command=self.apply_folder_changes, width=25).pack()

    def _show_main_menu(self):
        self.folders_frame.pack_forget()
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _show_folders_screen(self):
        self.main_canvas.pack_forget()
        self.main_scrollbar.pack_forget()
        self.folders_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._check_folders_state()

    def _select_all(self):
        for v in self.checkbox_vars.values():
            v.set(True)

    def _deselect_all(self):
        for v in self.checkbox_vars.values():
            v.set(False)

    # ---- initial state checks ----
    def _check_initial_states(self):
        self._check_bitlocker_state()
        self._check_photo_viewer_state()
        self._check_file_ext_state()
        self._check_recent_folders_state()
        self._check_image_details_state()

    def _check_folders_state(self):
        for guid in FOLDERS:
            r = _run_reg('query', f"{NAMESPACE_KEY}\\{guid}")
            ok = r.returncode == 0
            self.checkbox_vars[guid].set(ok)
            self.status_labels[guid].config(
                text="✓" if ok else "✗",
                foreground='#27ae60' if ok else '#e74c3c')

    def _check_bitlocker_state(self):
        """BitLocker скрыт, если ProgrammaticAccessOnly есть в ключе encrypt-bde.
        Ключ encrypt-bde-elev — вспомогательный (для запроса UAC), его наличие
        не обязательно для определения видимости пункта в меню."""
        try:
            r1 = _run_reg('query', r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde',
                          '/v', 'ProgrammaticAccessOnly')
            if r1.returncode == 0:
                self.bitlocker_enabled = False
                self.bitlocker_status.config(
                    text="Состояние: отключен (не отображается в меню)",
                    foreground='#e74c3c')
                self.bitlocker_btn.config(text="Включить BitLocker в меню")
            else:
                self.bitlocker_enabled = True
                self.bitlocker_status.config(
                    text="Состояние: включен (отображается в меню)",
                    foreground='#27ae60')
                self.bitlocker_btn.config(text="Отключить BitLocker из меню")
        except Exception:
            self.bitlocker_status.config(text="Состояние: ошибка проверки",
                                         foreground='#e67e22')
            self.bitlocker_btn.config(text="Ошибка проверки")

    def _check_photo_viewer_state(self):
        try:
            r1 = _run_reg('query',
                          r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
                          '/v', '.jpg')
            if r1.returncode == 0:
                r2 = _run_reg('query',
                              r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg')
                if r2.returncode == 0:
                    self.photo_viewer_enabled = True
                    self.photo_viewer_status.config(
                        text="Состояние: включен (форматы: .jpg, .png, .gif и др.)",
                        foreground='#27ae60')
                    self.photo_viewer_btn.config(text="Отключить Windows Photo Viewer")
                else:
                    self.photo_viewer_enabled = False
                    self.photo_viewer_status.config(
                        text="Состояние: частично установлен",
                        foreground='#e67e22')
                    self.photo_viewer_btn.config(
                        text="Включить Windows Photo Viewer полностью")
            else:
                self.photo_viewer_enabled = False
                self.photo_viewer_status.config(
                    text="Состояние: отключен (используется 'Фотографии')",
                    foreground='#e74c3c')
                self.photo_viewer_btn.config(text="Включить Windows Photo Viewer")
        except Exception:
            self.photo_viewer_status.config(text="Состояние: ошибка проверки",
                                            foreground='#e67e22')
            self.photo_viewer_btn.config(text="Ошибка проверки")

    def _check_file_ext_state(self):
        try:
            r = _run_reg('query',
                         r'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
                         '/v', 'HideFileExt')
            if r.returncode == 0:
                for line in r.stdout.split('\n'):
                    if 'HideFileExt' in line:
                        value = line.split('REG_DWORD')[1].strip()
                        self.file_ext_enabled = value == '0x0'
                        self.file_ext_status.config(
                            text="Состояние: включен (расширения отображаются)"
                            if self.file_ext_enabled
                            else "Состояние: отключен (расширения скрыты)",
                            foreground='#27ae60' if self.file_ext_enabled else '#e74c3c')
                        self.file_ext_btn.config(
                            text="Скрыть расширения файлов" if self.file_ext_enabled
                            else "Показать расширения файлов")
                        return
            self.file_ext_enabled = False
            self.file_ext_status.config(text="Состояние: отключен (расширения скрыты)",
                                        foreground='#e74c3c')
            self.file_ext_btn.config(text="Показать расширения файлов")
        except Exception:
            self.file_ext_status.config(text="Состояние: ошибка проверки",
                                        foreground='#e67e22')
            self.file_ext_btn.config(text="Ошибка проверки")

    def _check_recent_folders_state(self):
        try:
            r = _run_reg('query',
                         r'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer',
                         '/v', 'NoRecentDocsHistory')
            if r.returncode == 0:
                for line in r.stdout.split('\n'):
                    if 'NoRecentDocsHistory' in line:
                        value = line.split('REG_DWORD')[1].strip()
                        self.recent_folders_enabled = value != '0x1'
                        self.recent_folders_status.config(
                            text="Состояние: отключен (недавние папки скрыты)"
                            if not self.recent_folders_enabled
                            else "Состояние: включен (недавние папки отображаются)",
                            foreground='#27ae60' if self.recent_folders_enabled else '#e74c3c')
                        self.recent_folders_btn.config(
                            text="Включить недавние папки"
                            if not self.recent_folders_enabled
                            else "Отключить недавние папки")
                        return
            self.recent_folders_enabled = True
            self.recent_folders_status.config(
                text="Состояние: включен (по умолчанию)", foreground='#27ae60')
            self.recent_folders_btn.config(text="Отключить недавние папки")
        except Exception:
            self.recent_folders_status.config(text="Состояние: ошибка проверки",
                                              foreground='#e67e22')
            self.recent_folders_btn.config(text="Ошибка проверки")

    def _check_image_details_state(self):
        if not self.image_details_supported:
            self.image_details_hidden = True
            self.image_details_status.config(
                text="Состояние: отключен (значок скрыт)", foreground='#e74c3c')
            self.image_details_btn.config(text="Включить значок")
            return
        try:
            r = _run_reg('query', IMAGE_DETAILS_KEY, '/v', IMAGE_DETAILS_VALUE)
            if r.returncode == 0 and IMAGE_DETAILS_VALUE in r.stdout:
                match = re.search(r'REG_DWORD\s+(0x[0-9a-fA-F]+)', r.stdout)
                if match and match.group(1).lower() == '0x1':
                    self.image_details_hidden = True
                    self.image_details_status.config(
                        text="Состояние: отключен (значок скрыт)", foreground='#e74c3c')
                    self.image_details_btn.config(text="Включить значок")
                    return
            self.image_details_hidden = False
            self.image_details_status.config(
                text="Состояние: включен (значок отображается)", foreground='#27ae60')
            self.image_details_btn.config(text="Отключить значок")
        except Exception:
            self.image_details_status.config(text="Состояние: ошибка проверки",
                                             foreground='#e67e22')
            self.image_details_btn.config(text="Включить значок")

    # ---- toggles ----
    def toggle_image_details(self):
        if not self.image_details_supported:
            return
        if self.image_details_hidden:
            _run_reg('delete', IMAGE_DETAILS_KEY, '/v', IMAGE_DETAILS_VALUE, '/f')
        else:
            _run_reg('add', IMAGE_DETAILS_KEY, '/v', IMAGE_DETAILS_VALUE,
                     '/t', 'REG_DWORD', '/d', '1', '/f')
        self.image_details_hidden = not self.image_details_hidden
        self._check_image_details_state()

    def toggle_bitlocker(self):
        keys = [r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde',
                r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde-elev']
        if self.bitlocker_enabled:
            # Сейчас показан → скрываем: добавляем ProgrammaticAccessOnly
            for key in keys:
                _run_reg('add', key, '/v', 'ProgrammaticAccessOnly',
                         '/t', 'REG_SZ', '/d', '', '/f')
        else:
            # Сейчас скрыт → показываем: удаляем ProgrammaticAccessOnly
            for key in keys:
                _run_reg('delete', key, '/v', 'ProgrammaticAccessOnly', '/f')
        self._check_bitlocker_state()

    def toggle_photo_viewer(self):
        if self.photo_viewer_enabled:
            self._disable_photo_viewer()
        else:
            self._enable_photo_viewer()
        self._check_photo_viewer_state()

    def toggle_file_ext(self):
        value = "1" if self.file_ext_enabled else "0"
        _run_reg('add',
                 r'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
                 '/v', 'HideFileExt', '/t', 'REG_DWORD', '/d', value, '/f')
        self._check_file_ext_state()

    def toggle_recent_folders(self):
        reg = r'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer'
        if self.recent_folders_enabled:
            _run_reg('add', reg, '/v', 'NoRecentDocsHistory',
                     '/t', 'REG_DWORD', '/d', '1', '/f')
        else:
            _run_reg('delete', reg, '/v', 'NoRecentDocsHistory', '/f')
        self._check_recent_folders_state()

    def _enable_photo_viewer(self):
        base = r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations'
        for ext, value in PHOTO_VIEWER_EXTS.items():
            _run_reg('add', base, '/v', ext, '/t', 'REG_SZ', '/d', value, '/f')

        for class_key in PHOTO_VIEWER_CLASSES:
            _run_reg('add', class_key, '/f')

        flags = {
            PHOTO_VIEWER_CLASSES[0]: [('ImageOptionFlags', '1')],
            PHOTO_VIEWER_CLASSES[1]: [('EditFlags', '65536'), ('ImageOptionFlags', '1')],
            PHOTO_VIEWER_CLASSES[2]: [('EditFlags', '65536'), ('ImageOptionFlags', '1')],
            PHOTO_VIEWER_CLASSES[3]: [('ImageOptionFlags', '1')],
            PHOTO_VIEWER_CLASSES[4]: [('EditFlags', '65536'), ('ImageOptionFlags', '1')],
        }
        for key, values in flags.items():
            for name, data in values:
                _run_reg('add', key, '/v', name, '/t', 'REG_DWORD', '/d', data, '/f')

        command = (r'%SystemRoot%\System32\rundll32.exe '
                   r'"%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll", '
                   r'ImageView_Fullscreen %1')
        for class_key in PHOTO_VIEWER_CLASSES:
            for sub in (r'\shell', r'\shell\open', r'\shell\open\command'):
                _run_reg('add', class_key + sub, '/f')
            _run_reg('add', class_key + r'\shell\open\command',
                     '/ve', '/t', 'REG_SZ', '/d', command, '/f')
            _run_reg('add', class_key + r'\shell\open\DropTarget',
                     '/v', 'Clsid', '/t', 'REG_SZ',
                     '/d', '{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}', '/f')

        icons = {
            PHOTO_VIEWER_CLASSES[0]: r'%SystemRoot%\System32\imageres.dll,-70',
            PHOTO_VIEWER_CLASSES[1]: r'%SystemRoot%\System32\imageres.dll,-72',
            PHOTO_VIEWER_CLASSES[2]: r'%SystemRoot%\System32\imageres.dll,-72',
            PHOTO_VIEWER_CLASSES[3]: r'%SystemRoot%\System32\imageres.dll,-71',
            PHOTO_VIEWER_CLASSES[4]: r'%SystemRoot%\System32\wmphoto.dll,-400',
        }
        for key, icon in icons.items():
            _run_reg('add', key + r'\DefaultIcon', '/ve',
                     '/t', 'REG_SZ', '/d', icon, '/f')

    def _disable_photo_viewer(self):
        for ext in PHOTO_VIEWER_EXTS:
            _run_reg('delete',
                     r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
                     '/v', ext, '/f')
        for class_key in PHOTO_VIEWER_CLASSES + [r'HKLM\SOFTWARE\Classes\Applications\photoviewer.dll']:
            _run_reg('delete', class_key, '/f')
        for ext in ('.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff',
                    '.wdp', '.jfif'):
            _run_reg('delete', fr'HKCU\SOFTWARE\Classes\{ext}', '/f')

    def apply_folder_changes(self):
        for guid, var in self.checkbox_vars.items():
            desired = var.get()
            reg_path = f"{NAMESPACE_KEY}\\{guid}"
            if desired:
                _run_reg('add', reg_path, '/f')
            else:
                _run_reg('delete', reg_path, '/f')
            self.status_labels[guid].config(
                text="✓" if desired else "✗",
                foreground='#27ae60' if desired else '#e74c3c')

    # ---- typed paths ----
    def _on_clear_history_shift_click(self, event):
        self._open_typed_paths_window()
        return "break"

    def clear_typed_paths(self):
        reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
        r = _run_reg('query', reg_path)
        if r.returncode != 0:
            messagebox.showinfo("Очистка истории", "История очищена")
            return
        names = []
        for line in r.stdout.splitlines():
            line = line.strip()
            if not line or line.startswith('HKEY'):
                continue
            parts = line.split()
            if len(parts) >= 2 and parts[1] in ('REG_SZ', 'REG_EXPAND_SZ', 'REG_DWORD'):
                names.append(parts[0])
        for name in names:
            _run_reg('delete', reg_path, '/v', name, '/f')
        messagebox.showinfo("Очистка истории", "История очищена")

    def _read_typed_paths(self):
        reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
        items = []
        r = _run_reg('query', reg_path)
        if r.returncode != 0:
            return items
        for line in r.stdout.splitlines():
            line = line.strip()
            if not line or line.startswith('HKEY'):
                continue
            parts = line.split(None, 2)
            if len(parts) >= 3 and parts[1] in ('REG_SZ', 'REG_EXPAND_SZ'):
                items.append((parts[0], parts[2]))
        return items

    def _open_typed_paths_window(self):
        win = tk.Toplevel(self.parent)
        win.title("Выборочная очистка истории")
        win.geometry("340x380")
        win.transient(self.parent)
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            if os.path.exists(icon_path):
                win.iconbitmap(icon_path)
        except Exception:
            pass

        ttk.Label(win, text="Выберите для удаления:",
                  font=('Arial', 10, 'bold')).pack(anchor=tk.W, padx=10, pady=(10, 5))

        list_frame = ttk.Frame(win)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        vars_list = []
        items = self._read_typed_paths()
        if not items:
            ttk.Label(inner, text="История пуста.", foreground='#7f8c8d').pack(
                anchor=tk.W, pady=5)
        else:
            for name, value in items:
                row = ttk.Frame(inner)
                row.pack(fill=tk.X, padx=5, pady=2)
                v = tk.BooleanVar(value=False)
                vars_list.append((name, v))
                ttk.Checkbutton(row, variable=v).pack(side=tk.LEFT)
                ttk.Label(row, text=value, font=('Arial', 9)).pack(
                    side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

        def delete_selected():
            reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
            for name, v in vars_list:
                if v.get():
                    _run_reg('delete', reg_path, '/v', name, '/f')
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btns, text="Удалить выбранные", command=delete_selected).pack(
            side=tk.LEFT)
        ttk.Button(btns, text="Закрыть", command=win.destroy).pack(side=tk.RIGHT)

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()