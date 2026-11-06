import subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import sys
import os


class ExplorerModule:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)

        self.folders = {
            "{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}": ["Видео", "Папка с видео файлами"],
            "{088e3905-0323-4b02-9826-5d99428e115f}": ["Загрузки", "Папка для загруженных файлов"],
            "{d3162b92-9365-467a-956b-92703aca08af}": ["Документы", "Папка для документов"],
            "{24ad3ad4-a569-4530-98e1-ab02f9417aa8}": ["Изображения", "Папка с изображениями"],
            "{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}": ["Музыка", "Папка с аудио файлами"],
            "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}": ["Рабочий стол", "Папка рабочего стола"],
            "{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}": ["3D объекты", "Папка для 3D моделей"]
        }

        self.checkbox_vars = {}
        self.bitlocker_enabled = False
        self.photo_viewer_enabled = False
        self.file_ext_enabled = False
        self.recent_folders_enabled = True
        self.image_details_hidden = False
        self.image_details_supported = False
        self.status_labels = {}
        self.current_screen = "main"

        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.main_container = tk.Frame(self.content_frame)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        self.main_canvas = tk.Canvas(self.main_container, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self.main_container, orient="vertical",
                                            command=self.main_canvas.yview)
        self.main_scrollable_frame = ttk.Frame(self.main_canvas)
        self.main_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self.main_canvas_window = self.main_canvas.create_window(
            (0, 0), window=self.main_scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        self.main_canvas.bind(
            "<Configure>",
            lambda e: self.main_canvas.itemconfigure(self.main_canvas_window, width=e.width))

        self.main_menu_frame = ttk.Frame(self.main_scrollable_frame)
        self.folders_frame = ttk.Frame(self.main_container)

        self.create_main_menu()
        self.create_folders_screen()

        self.show_main_menu()
        self.check_initial_states()
        self.show()

    def _is_windows_10_22h2_or_higher(self):
        try:
            result = subprocess.run(
                ['reg', 'query', r'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion',
                 '/v', 'CurrentBuildNumber'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                for line in result.stdout.splitlines():
                    if 'CurrentBuildNumber' in line:
                        build = int(line.split('REG_SZ')[1].strip())
                        return 19045 <= build < 22000
        except Exception:
            pass
        try:
            ver = sys.getwindowsversion()
            if ver.major == 10 and 19045 <= ver.build < 22000:
                return True
            return False
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

    def create_main_menu(self):
        self.main_menu_frame.pack(fill=tk.BOTH, expand=True, padx=(20, 35), pady=10)

        ttk.Label(self.main_menu_frame, text="Настройки Проводника",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=(0, 30))

        ttk.Button(self.main_menu_frame, text="Папки в Этот компьютер",
                   command=self.show_folders_screen, width=35).pack(pady=10)

        clear_history_btn = ttk.Button(self.main_menu_frame, text="Очистка истории",
                                       command=self.clear_typed_paths, width=35)
        clear_history_btn.pack(pady=(0, 10))
        self._add_tooltip(clear_history_btn, "Для выборочного удаления зажмите Shift")
        clear_history_btn.bind("<Shift-Button-1>", self._on_clear_history_shift_click)

        ttk.Separator(self.main_menu_frame, orient='horizontal').pack(fill=tk.X, pady=20)

        # BitLocker
        bitlocker_frame = ttk.Frame(self.main_menu_frame)
        bitlocker_frame.pack(fill=tk.X, pady=5)
        ttk.Label(bitlocker_frame,
                  text="Управление пунктом BitLocker в контекстном меню дисков:",
                  font=('Arial', 9), foreground='#7f8c8d').pack(anchor=tk.W)
        self.bitlocker_status = ttk.Label(bitlocker_frame, text="Состояние: проверяется...",
                                          font=('Arial', 9), foreground='#95a5a6')
        self.bitlocker_status.pack(anchor=tk.W, pady=(2, 5))
        self.bitlocker_btn = ttk.Button(bitlocker_frame, text="---",
                                        command=self.toggle_bitlocker, width=35)
        self.bitlocker_btn.pack()

        ttk.Separator(self.main_menu_frame, orient='horizontal').pack(fill=tk.X, pady=20)

        # Photo Viewer
        pv_frame = ttk.Frame(self.main_menu_frame)
        pv_frame.pack(fill=tk.X, pady=5)
        ttk.Label(pv_frame, text="Включение/отключение Windows Photo Viewer:",
                  font=('Arial', 9), foreground='#7f8c8d').pack(anchor=tk.W)
        self.photo_viewer_status = ttk.Label(pv_frame, text="Состояние: проверяется...",
                                             font=('Arial', 9), foreground='#95a5a6')
        self.photo_viewer_status.pack(anchor=tk.W, pady=(2, 5))
        self.photo_viewer_btn = ttk.Button(pv_frame, text="---",
                                           command=self.toggle_photo_viewer, width=35)
        self.photo_viewer_btn.pack()

        ttk.Separator(self.main_menu_frame, orient='horizontal').pack(fill=tk.X, pady=20)

        # File extensions
        fe_frame = ttk.Frame(self.main_menu_frame)
        fe_frame.pack(fill=tk.X, pady=5)
        ttk.Label(fe_frame, text="Показ расширений файлов в Проводнике:",
                  font=('Arial', 9), foreground='#7f8c8d').pack(anchor=tk.W)
        self.file_ext_status = ttk.Label(fe_frame, text="Состояние: проверяется...",
                                         font=('Arial', 9), foreground='#95a5a6')
        self.file_ext_status.pack(anchor=tk.W, pady=(2, 5))
        self.file_ext_btn = ttk.Button(fe_frame, text="---",
                                       command=self.toggle_file_ext, width=35)
        self.file_ext_btn.pack()

        ttk.Separator(self.main_menu_frame, orient='horizontal').pack(fill=tk.X, pady=20)

        # Recent folders
        rf_frame = ttk.Frame(self.main_menu_frame)
        rf_frame.pack(fill=tk.X, pady=5)
        ttk.Label(rf_frame, text="Отображение недавних папок в Проводнике:",
                  font=('Arial', 9), foreground='#7f8c8d').pack(anchor=tk.W)
        self.recent_folders_status = ttk.Label(rf_frame, text="Состояние: проверяется...",
                                               font=('Arial', 9), foreground='#95a5a6')
        self.recent_folders_status.pack(anchor=tk.W, pady=(2, 5))
        self.recent_folders_btn = ttk.Button(rf_frame, text="---",
                                             command=self.toggle_recent_folders, width=35)
        self.recent_folders_btn.pack()

        ttk.Separator(self.main_menu_frame, orient='horizontal').pack(fill=tk.X, pady=20)

        # Image details
        id_frame = ttk.Frame(self.main_menu_frame)
        id_frame.pack(fill=tk.X, pady=5)
        ttk.Label(id_frame,
                  text="Отображение Дополнительные сведения об этом изображении",
                  font=('Arial', 9), foreground='#7f8c8d').pack(anchor=tk.W)
        self.image_details_status = ttk.Label(id_frame, text="Состояние: проверяется...",
                                              font=('Arial', 9), foreground='#95a5a6')
        self.image_details_status.pack(anchor=tk.W, pady=(2, 5))
        self.image_details_btn = ttk.Button(id_frame, text="Включить значок",
                                            command=self.toggle_image_details, width=35)
        self.image_details_btn.pack()

        self.image_details_supported = self._is_windows_10_22h2_or_higher()
        if not self.image_details_supported:
            self.image_details_btn.config(state=tk.DISABLED, text="Включить значок")
            self._add_tooltip(self.image_details_btn,
                              "Эта функция доступна только в Windows 10 22H2 и выше")

    def _add_tooltip(self, widget, text):
        tooltip = tk.Toplevel(widget)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_geometry("+0+0")
        tooltip.withdraw()
        tk.Label(tooltip, text=text, background="#ffffe0", relief="solid",
                 borderwidth=1, font=("Arial", 9)).pack()

        def enter(event):
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 5
            tooltip.wm_geometry(f"+{x}+{y}")
            tooltip.deiconify()

        def leave(event):
            tooltip.withdraw()

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def _on_clear_history_shift_click(self, event):
        self.open_typed_paths_window()
        return "break"

    def clear_typed_paths(self):
        reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
        try:
            result = subprocess.run(['reg', 'query', reg_path],
                                    capture_output=True, text=True, encoding='cp866',
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode != 0:
                messagebox.showinfo("Очистка истории", "История очищена")
                return
            value_names = []
            for line in result.stdout.splitlines():
                line = line.strip()
                if not line or line.startswith('HKEY'):
                    continue
                parts = line.split()
                if len(parts) >= 2 and parts[1] in ('REG_SZ', 'REG_EXPAND_SZ', 'REG_DWORD'):
                    value_names.append(parts[0])

            for name in value_names:
                subprocess.run(['reg', 'delete', reg_path, '/v', name, '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
            messagebox.showinfo("Очистка истории", "История очищена")
        except Exception as e:
            print(f"Ошибка при очистке истории TypedPaths: {e}")

    def _read_typed_paths(self):
        reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
        items = []
        try:
            result = subprocess.run(['reg', 'query', reg_path],
                                    capture_output=True, text=True, encoding='cp866',
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode != 0:
                return items
            for line in result.stdout.splitlines():
                line = line.strip()
                if not line or line.startswith('HKEY'):
                    continue
                parts = line.split(None, 2)
                if len(parts) >= 3 and parts[1] in ('REG_SZ', 'REG_EXPAND_SZ'):
                    items.append((parts[0], parts[2]))
        except Exception as e:
            print(f"Ошибка при чтении TypedPaths: {e}")
        return items

    def open_typed_paths_window(self):
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

        var_list = []
        items = self._read_typed_paths()

        if not items:
            ttk.Label(inner, text="История пуста.", foreground='#7f8c8d').pack(
                anchor=tk.W, pady=5)
        else:
            for name, value in items:
                row = ttk.Frame(inner)
                row.pack(fill=tk.X, padx=5, pady=2)
                var = tk.BooleanVar(value=False)
                var_list.append((name, var))
                ttk.Checkbutton(row, variable=var).pack(side=tk.LEFT)
                ttk.Label(row, text=value, font=('Arial', 9)).pack(
                    side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

        def delete_selected():
            reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
            for name, var in var_list:
                if var.get():
                    try:
                        subprocess.run(['reg', 'delete', reg_path, '/v', name, '/f'],
                                       capture_output=True, text=True, encoding='cp866',
                                       creationflags=subprocess.CREATE_NO_WINDOW)
                    except Exception:
                        pass
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btn_frame, text="Удалить выбранные",
                   command=delete_selected).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Закрыть", command=win.destroy).pack(side=tk.RIGHT)

    def create_folders_screen(self):
        header_frame = ttk.Frame(self.folders_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Button(header_frame, text="← Назад",
                   command=self.show_main_menu, width=10).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="Папки в Этот компьютер",
                  font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(side=tk.LEFT, padx=20)

        ttk.Label(self.folders_frame,
                  text="Выберите какие папки отображать в Этот компьютер",
                  font=('Arial', 10), foreground='#7f8c8d').pack(pady=(0, 15))

        folders_list_frame = ttk.LabelFrame(self.folders_frame,
                                            text="Доступные папки", padding=15)
        folders_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        canvas = tk.Canvas(folders_list_frame, highlightthickness=0, height=220)
        scrollbar = ttk.Scrollbar(folders_list_frame, orient="vertical",
                                  command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for guid, (name, description) in self.folders.items():
            frame = ttk.Frame(self.scrollable_frame)
            frame.pack(fill=tk.X, padx=5, pady=4)
            var = tk.BooleanVar()
            self.checkbox_vars[guid] = var
            ttk.Checkbutton(frame, text=name, variable=var,
                            onvalue=True, offvalue=False).pack(side=tk.LEFT, anchor=tk.W)

            status_label = ttk.Label(frame, text="•", font=('Arial', 9, 'bold'), width=2)
            status_label.pack(side=tk.LEFT, padx=(2, 0))
            self.status_labels[guid] = status_label

            ttk.Label(frame, text=f" - {description}", font=('Arial', 9),
                      foreground='#95a5a6').pack(side=tk.LEFT, padx=(0, 8))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        top_buttons = ttk.Frame(self.folders_frame)
        top_buttons.pack(pady=15)
        ttk.Button(top_buttons, text="Выбрать все папки",
                   command=self.select_all, width=20).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(top_buttons, text="Снять все папки",
                   command=self.deselect_all, width=20).pack(side=tk.LEFT)

        bottom_buttons = ttk.Frame(self.folders_frame)
        bottom_buttons.pack(pady=(0, 15))
        ttk.Button(bottom_buttons, text="Применить изменения",
                   command=self.apply_folder_changes, width=25).pack()

    def check_initial_states(self):
        self.check_bitlocker_state()
        self.check_photo_viewer_state()
        self.check_file_ext_state()
        self.check_recent_folders_state()
        self.check_image_details_state()

    def show_main_menu(self):
        self.folders_frame.pack_forget()
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.current_screen = "main"

    def show_folders_screen(self):
        self.main_canvas.pack_forget()
        self.main_scrollbar.pack_forget()
        self.folders_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.check_current_state()
        self.current_screen = "folders"

    def check_current_state(self):
        base_key = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"
        for guid in self.folders.keys():
            try:
                result = subprocess.run(
                    ['reg', 'query', f"{base_key}\\{guid}"],
                    capture_output=True, text=True, encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    self.checkbox_vars[guid].set(True)
                    self.status_labels[guid].config(text="✓", foreground='#27ae60')
                else:
                    self.checkbox_vars[guid].set(False)
                    self.status_labels[guid].config(text="✗", foreground='#e74c3c')
            except Exception:
                self.checkbox_vars[guid].set(False)
                self.status_labels[guid].config(text="✗", foreground='#e74c3c')

    def check_bitlocker_state(self):
        try:
            r1 = subprocess.run(
                ['reg', 'query', r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde',
                 '/v', 'ProgrammaticAccessOnly'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            r2 = subprocess.run(
                ['reg', 'query', r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde-elev',
                 '/v', 'ProgrammaticAccessOnly'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if r1.returncode == 0 and r2.returncode == 0:
                self.bitlocker_enabled = False
                self.bitlocker_status.config(
                    text="Состояние: отключен (не отображается в меню)", foreground='#e74c3c')
                self.bitlocker_btn.config(text="Включить BitLocker в меню")
            else:
                self.bitlocker_enabled = True
                self.bitlocker_status.config(
                    text="Состояние: включен (отображается в меню)", foreground='#27ae60')
                self.bitlocker_btn.config(text="Отключить BitLocker из меню")
        except Exception:
            self.bitlocker_status.config(text="Состояние: ошибка проверки",
                                         foreground='#e67e22')
            self.bitlocker_btn.config(text="Ошибка проверки")

    def check_photo_viewer_state(self):
        try:
            r1 = subprocess.run(
                ['reg', 'query',
                 r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
                 '/v', '.jpg'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if r1.returncode == 0:
                r2 = subprocess.run(
                    ['reg', 'query', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg'],
                    capture_output=True, text=True, encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW)
                if r2.returncode == 0:
                    self.photo_viewer_enabled = True
                    self.photo_viewer_status.config(
                        text="Состояние: включен (форматы: .jpg, .png, .gif и др.)",
                        foreground='#27ae60')
                    self.photo_viewer_btn.config(text="Отключить Windows Photo Viewer")
                else:
                    self.photo_viewer_enabled = False
                    self.photo_viewer_status.config(
                        text="Состояние: частично установлен (требуется полное включение)",
                        foreground='#e67e22')
                    self.photo_viewer_btn.config(
                        text="Включить Windows Photo Viewer полностью")
            else:
                self.photo_viewer_enabled = False
                self.photo_viewer_status.config(
                    text="Состояние: отключен (используется приложение 'Фотографии')",
                    foreground='#e74c3c')
                self.photo_viewer_btn.config(text="Включить Windows Photo Viewer")
        except Exception:
            self.photo_viewer_status.config(text="Состояние: ошибка проверки",
                                            foreground='#e67e22')
            self.photo_viewer_btn.config(text="Ошибка проверки")

    def check_file_ext_state(self):
        try:
            result = subprocess.run(
                ['reg', 'query',
                 r'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
                 '/v', 'HideFileExt'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if 'HideFileExt' in line:
                        value = line.split('REG_DWORD')[1].strip()
                        if value == '0x0':
                            self.file_ext_enabled = True
                            self.file_ext_status.config(
                                text="Состояние: включен (расширения отображаются)",
                                foreground='#27ae60')
                            self.file_ext_btn.config(text="Скрыть расширения файлов")
                        else:
                            self.file_ext_enabled = False
                            self.file_ext_status.config(
                                text="Состояние: отключен (расширения скрыты)",
                                foreground='#e74c3c')
                            self.file_ext_btn.config(text="Показать расширения файлов")
                        break
            else:
                self.file_ext_enabled = False
                self.file_ext_status.config(
                    text="Состояние: отключен (расширения скрыты, по умолчанию)",
                    foreground='#e74c3c')
                self.file_ext_btn.config(text="Показать расширения файлов")
        except Exception:
            self.file_ext_status.config(text="Состояние: ошибка проверки",
                                        foreground='#e67e22')
            self.file_ext_btn.config(text="Ошибка проверки")

    def check_recent_folders_state(self):
        try:
            result = subprocess.run(
                ['reg', 'query',
                 r'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer',
                 '/v', 'NoRecentDocsHistory'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if 'NoRecentDocsHistory' in line:
                        value = line.split('REG_DWORD')[1].strip()
                        if value == '0x1':
                            self.recent_folders_enabled = False
                            self.recent_folders_status.config(
                                text="Состояние: отключен (недавние папки скрыты)",
                                foreground='#e74c3c')
                            self.recent_folders_btn.config(text="Включить недавние папки")
                        else:
                            self.recent_folders_enabled = True
                            self.recent_folders_status.config(
                                text="Состояние: включен (недавние папки отображаются)",
                                foreground='#27ae60')
                            self.recent_folders_btn.config(text="Отключить недавние папки")
                        break
            else:
                self.recent_folders_enabled = True
                self.recent_folders_status.config(
                    text="Состояние: включен (параметр отсутствует, по умолчанию)",
                    foreground='#27ae60')
                self.recent_folders_btn.config(text="Отключить недавние папки")
        except Exception:
            self.recent_folders_status.config(text="Состояние: ошибка проверки",
                                              foreground='#e67e22')
            self.recent_folders_btn.config(text="Ошибка проверки")

    def check_image_details_state(self):
        if not self.image_details_supported:
            self.image_details_hidden = True
            self.image_details_status.config(
                text="Состояние: отключен (значок скрыт)", foreground='#e74c3c')
            self.image_details_btn.config(text="Включить значок")
            return

        reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        value_name = "{2cc5ca98-6485-489a-920e-b3e88a6ccce3}"
        try:
            result = subprocess.run(['reg', 'query', reg_path, '/v', value_name],
                                    capture_output=True, text=True, encoding='cp866',
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0 and value_name in result.stdout:
                import re
                match = re.search(r'REG_DWORD\s+(0x[0-9a-fA-F]+)', result.stdout)
                if match:
                    if match.group(1).lower() == '0x1':
                        self.image_details_hidden = True
                        self.image_details_status.config(
                            text="Состояние: отключен (значок скрыт)", foreground='#e74c3c')
                        self.image_details_btn.config(text="Включить значок")
                    else:
                        self.image_details_hidden = False
                        self.image_details_status.config(
                            text="Состояние: включен (значок отображается)",
                            foreground='#27ae60')
                        self.image_details_btn.config(text="Отключить значок")
                else:
                    self.image_details_hidden = False
                    self.image_details_status.config(
                        text="Состояние: включен (значок отображается)",
                        foreground='#27ae60')
                    self.image_details_btn.config(text="Отключить значок")
            else:
                self.image_details_hidden = False
                self.image_details_status.config(
                    text="Состояние: включен (значок отображается)",
                    foreground='#27ae60')
                self.image_details_btn.config(text="Отключить значок")
        except Exception:
            self.image_details_status.config(text="Состояние: ошибка проверки",
                                             foreground='#e67e22')
            self.image_details_btn.config(text="Включить значок")

    def toggle_image_details(self):
        if not self.image_details_supported:
            return
        reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        value_name = "{2cc5ca98-6485-489a-920e-b3e88a6ccce3}"
        try:
            if self.image_details_hidden:
                subprocess.run(['reg', 'delete', reg_path, '/v', value_name, '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                subprocess.run(['reg', 'add', reg_path, '/v', value_name,
                                '/t', 'REG_DWORD', '/d', '1', '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
            self.image_details_hidden = not self.image_details_hidden
            self.check_image_details_state()
        except Exception as e:
            print(f"Ошибка при переключении видимости значка: {e}")

    def apply_folder_changes(self):
        base_key = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"
        for guid, var in self.checkbox_vars.items():
            desired_state = var.get()
            reg_path = f"{base_key}\\{guid}"
            try:
                if desired_state:
                    subprocess.run(['reg', 'add', reg_path, '/f'], capture_output=True,
                                   text=True, encoding='cp866',
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                    self.status_labels[guid].config(text="✓", foreground='#27ae60')
                else:
                    subprocess.run(['reg', 'delete', reg_path, '/f'], capture_output=True,
                                   text=True, encoding='cp866',
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                    self.status_labels[guid].config(text="✗", foreground='#e74c3c')
            except Exception:
                self.status_labels[guid].config(
                    text="✓" if desired_state else "✗",
                    foreground='#27ae60' if desired_state else '#e74c3c')

    def toggle_bitlocker(self):
        try:
            action = "disable" if self.bitlocker_enabled else "enable"
            for key in [r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde',
                        r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde-elev']:
                try:
                    if action == "disable":
                        subprocess.run(['reg', 'add', key, '/v', 'ProgrammaticAccessOnly',
                                        '/t', 'REG_SZ', '/d', '', '/f'],
                                       capture_output=True, text=True, encoding='cp866',
                                       creationflags=subprocess.CREATE_NO_WINDOW)
                    else:
                        subprocess.run(['reg', 'delete', key, '/v',
                                        'ProgrammaticAccessOnly', '/f'],
                                       capture_output=True, text=True, encoding='cp866',
                                       creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception:
                    pass
            self.bitlocker_enabled = not self.bitlocker_enabled
            self.check_bitlocker_state()
        except Exception:
            pass

    def toggle_photo_viewer(self):
        try:
            if self.photo_viewer_enabled:
                self.disable_photo_viewer()
            else:
                self.enable_photo_viewer()
            self.photo_viewer_enabled = not self.photo_viewer_enabled
            self.check_photo_viewer_state()
        except Exception as e:
            print(f"Ошибка при переключении Photo Viewer: {e}")

    def toggle_file_ext(self):
        try:
            value = "1" if self.file_ext_enabled else "0"
            reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            subprocess.run(['reg', 'add', reg_path, '/v', 'HideFileExt',
                            '/t', 'REG_DWORD', '/d', value, '/f'],
                           capture_output=True, text=True, encoding='cp866',
                           creationflags=subprocess.CREATE_NO_WINDOW)
            self.file_ext_enabled = not self.file_ext_enabled
            self.check_file_ext_state()
        except Exception as e:
            print(f"Ошибка при переключении показа расширений: {e}")

    def toggle_recent_folders(self):
        try:
            reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
            if self.recent_folders_enabled:
                subprocess.run(['reg', 'add', reg_path, '/v', 'NoRecentDocsHistory',
                                '/t', 'REG_DWORD', '/d', '1', '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                subprocess.run(['reg', 'delete', reg_path, '/v',
                                'NoRecentDocsHistory', '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
            self.recent_folders_enabled = not self.recent_folders_enabled
            self.check_recent_folders_state()
        except Exception as e:
            print(f"Ошибка при переключении показа недавних папок: {e}")

    def enable_photo_viewer(self):
        try:
            reg_keys = [
                ('.jpg', 'PhotoViewer.FileAssoc.Jpeg'),
                ('.wdp', 'PhotoViewer.FileAssoc.Wdp'),
                ('.jfif', 'PhotoViewer.FileAssoc.JFIF'),
                ('.dib', 'PhotoViewer.FileAssoc.Bitmap'),
                ('.png', 'PhotoViewer.FileAssoc.Png'),
                ('.tiff', 'PhotoViewer.FileAssoc.Tiff'),
                ('.jxr', 'PhotoViewer.FileAssoc.Wdp'),
                ('.bmp', 'PhotoViewer.FileAssoc.Bitmap'),
                ('.jpe', 'PhotoViewer.FileAssoc.Jpeg'),
                ('.jpeg', 'PhotoViewer.FileAssoc.Jpeg'),
                ('.gif', 'PhotoViewer.FileAssoc.Gif'),
                ('.tif', 'PhotoViewer.FileAssoc.Tiff'),
            ]
            base = r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations'
            for ext, value in reg_keys:
                subprocess.run(['reg', 'add', base, '/v', ext, '/t', 'REG_SZ', '/d', value, '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)

            main_classes = [
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Gif',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Tiff'
            ]
            for class_key in main_classes:
                subprocess.run(['reg', 'add', class_key, '/f'], capture_output=True,
                               text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)

            flags = {
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap': [('ImageOptionFlags', '1')],
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF': [
                    ('EditFlags', '65536'), ('ImageOptionFlags', '1')],
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg': [
                    ('EditFlags', '65536'), ('ImageOptionFlags', '1')],
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png': [('ImageOptionFlags', '1')],
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp': [
                    ('EditFlags', '65536'), ('ImageOptionFlags', '1')],
            }
            for key, values in flags.items():
                for name, data in values:
                    subprocess.run(['reg', 'add', key, '/v', name, '/t', 'REG_DWORD',
                                    '/d', data, '/f'],
                                   capture_output=True, text=True, encoding='cp866',
                                   creationflags=subprocess.CREATE_NO_WINDOW)

            command_value = r'%SystemRoot%\System32\rundll32.exe "%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll", ImageView_Fullscreen %1'
            for class_key in main_classes:
                for sub in [r'\shell', r'\shell\open', r'\shell\open\command']:
                    subprocess.run(['reg', 'add', class_key + sub, '/f'], capture_output=True,
                                   text=True, encoding='cp866',
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(['reg', 'add', class_key + r'\shell\open\command',
                                '/ve', '/t', 'REG_SZ', '/d', command_value, '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(['reg', 'add', class_key + r'\shell\open\DropTarget',
                                '/v', 'Clsid', '/t', 'REG_SZ',
                                '/d', '{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}', '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)

            icons = {
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap':
                    '%SystemRoot%\\System32\\imageres.dll,-70',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF':
                    '%SystemRoot%\\System32\\imageres.dll,-72',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg':
                    '%SystemRoot%\\System32\\imageres.dll,-72',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png':
                    '%SystemRoot%\\System32\\imageres.dll,-71',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp':
                    '%SystemRoot%\\System32\\wmphoto.dll,-400',
            }
            for key, icon in icons.items():
                subprocess.run(['reg', 'add', key + r'\DefaultIcon', '/ve',
                                '/t', 'REG_SZ', '/d', icon, '/f'],
                               capture_output=True, text=True, encoding='cp866',
                               creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e:
            print(f"Ошибка при включении Photo Viewer: {e}")

    def disable_photo_viewer(self):
        try:
            for ext in ['.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff',
                        '.wdp', '.jfif', '.jpe', '.dib', '.jxr']:
                try:
                    subprocess.run(
                        ['reg', 'delete',
                         r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
                         '/v', ext, '/f'],
                        capture_output=True, text=True, encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception:
                    pass

            for class_key in [
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Gif',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Tiff',
                r'HKLM\SOFTWARE\Classes\Applications\photoviewer.dll'
            ]:
                try:
                    subprocess.run(['reg', 'delete', class_key, '/f'], capture_output=True,
                                   text=True, encoding='cp866',
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception:
                    pass

            for ext in ['.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff',
                        '.wdp', '.jfif']:
                try:
                    subprocess.run(['reg', 'delete', fr'HKCU\SOFTWARE\Classes\{ext}', '/f'],
                                   capture_output=True, text=True, encoding='cp866',
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception:
                    pass
        except Exception as e:
            print(f"Ошибка при отключении Photo Viewer: {e}")

    def select_all(self):
        for var in self.checkbox_vars.values():
            var.set(True)

    def deselect_all(self):
        for var in self.checkbox_vars.values():
            var.set(False)

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()