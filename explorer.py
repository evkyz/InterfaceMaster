import subprocess
import tkinter as tk
from tkinter import ttk


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
        self.status_labels = {}
        self.current_screen = "main"

        # Основной контейнер
        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Главный контейнер с прокруткой
        self.main_container = tk.Frame(self.content_frame)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # Канвас и скроллбар для главного меню
        self.main_canvas = tk.Canvas(self.main_container, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(self.main_container, orient="vertical", command=self.main_canvas.yview)
        self.main_scrollable_frame = ttk.Frame(self.main_canvas)

        self.main_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        )

        self.main_canvas.create_window((0, 0), window=self.main_scrollable_frame, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        # Фрейм для главного меню внутри прокручиваемой области
        self.main_menu_frame = ttk.Frame(self.main_scrollable_frame)

        # Фрейм для управления папками
        self.folders_frame = ttk.Frame(self.main_container)

        self.create_main_menu()
        self.create_folders_screen()

        self.show_main_menu()
        self.check_initial_states()
        self.show()

    def create_main_menu(self):
        """Создание главного меню с кнопками"""
        self.main_menu_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        title_label = ttk.Label(
            self.main_menu_frame,
            text="Настройки Проводника",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=(0, 30))

        # Кнопка 1: Папки в Этот компьютер
        folders_btn = ttk.Button(
            self.main_menu_frame,
            text="Папки в Этот компьютер",
            command=self.show_folders_screen,
            width=35
        )
        folders_btn.pack(pady=10)

        # Разделитель
        separator = ttk.Separator(self.main_menu_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=20)

        # Кнопка 2: Включить/Отключить BitLocker
        bitlocker_frame = ttk.Frame(self.main_menu_frame)
        bitlocker_frame.pack(fill=tk.X, pady=5)

        bitlocker_desc = ttk.Label(
            bitlocker_frame,
            text="Управление пунктом BitLocker в контекстном меню дисков:",
            font=('Arial', 9),
            foreground='#7f8c8d'
        )
        bitlocker_desc.pack(anchor=tk.W)

        self.bitlocker_status = ttk.Label(
            bitlocker_frame,
            text="Состояние: проверяется...",
            font=('Arial', 9),
            foreground='#95a5a6'
        )
        self.bitlocker_status.pack(anchor=tk.W, pady=(2, 5))

        self.bitlocker_btn = ttk.Button(
            bitlocker_frame,
            text="---",
            command=self.toggle_bitlocker,
            width=35
        )
        self.bitlocker_btn.pack()

        # Разделитель
        separator2 = ttk.Separator(self.main_menu_frame, orient='horizontal')
        separator2.pack(fill=tk.X, pady=20)

        # Кнопка 3: Включить/Отключить Просмотр фотографий
        photo_viewer_frame = ttk.Frame(self.main_menu_frame)
        photo_viewer_frame.pack(fill=tk.X, pady=5)

        photo_viewer_desc = ttk.Label(
            photo_viewer_frame,
            text="Включение/отключение Windows Photo Viewer:",
            font=('Arial', 9),
            foreground='#7f8c8d'
        )
        photo_viewer_desc.pack(anchor=tk.W)

        self.photo_viewer_status = ttk.Label(
            photo_viewer_frame,
            text="Состояние: проверяется...",
            font=('Arial', 9),
            foreground='#95a5a6'
        )
        self.photo_viewer_status.pack(anchor=tk.W, pady=(2, 5))

        self.photo_viewer_btn = ttk.Button(
            photo_viewer_frame,
            text="---",
            command=self.toggle_photo_viewer,
            width=35
        )
        self.photo_viewer_btn.pack()

        # Разделитель
        separator3 = ttk.Separator(self.main_menu_frame, orient='horizontal')
        separator3.pack(fill=tk.X, pady=20)

        # Кнопка 4: Включить/Отключить показ расширений файлов
        file_ext_frame = ttk.Frame(self.main_menu_frame)
        file_ext_frame.pack(fill=tk.X, pady=5)

        file_ext_desc = ttk.Label(
            file_ext_frame,
            text="Показ расширений файлов в Проводнике:",
            font=('Arial', 9),
            foreground='#7f8c8d'
        )
        file_ext_desc.pack(anchor=tk.W)

        self.file_ext_status = ttk.Label(
            file_ext_frame,
            text="Состояние: проверяется...",
            font=('Arial', 9),
            foreground='#95a5a6'
        )
        self.file_ext_status.pack(anchor=tk.W, pady=(2, 5))

        self.file_ext_btn = ttk.Button(
            file_ext_frame,
            text="---",
            command=self.toggle_file_ext,
            width=35
        )
        self.file_ext_btn.pack()

        # Разделитель
        separator4 = ttk.Separator(self.main_menu_frame, orient='horizontal')
        separator4.pack(fill=tk.X, pady=20)

        # Кнопка 5: Включить/Отключить Недавние папки
        recent_folders_frame = ttk.Frame(self.main_menu_frame)
        recent_folders_frame.pack(fill=tk.X, pady=5)

        recent_folders_desc = ttk.Label(
            recent_folders_frame,
            text="Отображение недавних папок в Проводнике:",
            font=('Arial', 9),
            foreground='#7f8c8d'
        )
        recent_folders_desc.pack(anchor=tk.W)

        self.recent_folders_status = ttk.Label(
            recent_folders_frame,
            text="Состояние: проверяется...",
            font=('Arial', 9),
            foreground='#95a5a6'
        )
        self.recent_folders_status.pack(anchor=tk.W, pady=(2, 5))

        self.recent_folders_btn = ttk.Button(
            recent_folders_frame,
            text="---",
            command=self.toggle_recent_folders,
            width=35
        )
        self.recent_folders_btn.pack()

    def create_folders_screen(self):
        """Создание экрана управления папками"""
        # Заголовок и кнопка назад
        header_frame = ttk.Frame(self.folders_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        back_btn = ttk.Button(
            header_frame,
            text="← Назад",
            command=self.show_main_menu,
            width=10
        )
        back_btn.pack(side=tk.LEFT)

        title_label = ttk.Label(
            header_frame,
            text="Папки в Этот компьютер",
            font=('Arial', 14, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(side=tk.LEFT, padx=20)

        # Описание
        desc_label = ttk.Label(
            self.folders_frame,
            text="Выберите какие папки отображать в Этот компьютер",
            font=('Arial', 10),
            foreground='#7f8c8d'
        )
        desc_label.pack(pady=(0, 15))

        # Фрейм со списком папок
        folders_list_frame = ttk.LabelFrame(self.folders_frame, text="Доступные папки", padding=15)
        folders_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        canvas = tk.Canvas(folders_list_frame, highlightthickness=0, height=220)
        scrollbar = ttk.Scrollbar(folders_list_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for guid, (name, description) in self.folders.items():
            frame = ttk.Frame(self.scrollable_frame)
            frame.pack(fill=tk.X, padx=5, pady=4)

            var = tk.BooleanVar()
            self.checkbox_vars[guid] = var

            checkbox = ttk.Checkbutton(
                frame,
                text=name,
                variable=var,
                onvalue=True,
                offvalue=False
            )
            checkbox.pack(side=tk.LEFT, anchor=tk.W)

            status_label = ttk.Label(
                frame,
                text="•",
                font=('Arial', 9, 'bold'),
                width=2
            )
            status_label.pack(side=tk.LEFT, padx=(2, 0))
            self.status_labels[guid] = status_label

            desc_label = ttk.Label(
                frame,
                text=f" - {description}",
                font=('Arial', 9),
                foreground='#95a5a6'
            )
            desc_label.pack(side=tk.LEFT, padx=(0, 8))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Кнопки управления - первая строка
        buttons_frame_top = ttk.Frame(self.folders_frame)
        buttons_frame_top.pack(pady=15)

        select_all_btn = ttk.Button(
            buttons_frame_top,
            text="Выбрать все папки",
            command=self.select_all,
            width=20
        )
        select_all_btn.pack(side=tk.LEFT, padx=(0, 10))

        deselect_all_btn = ttk.Button(
            buttons_frame_top,
            text="Снять все папки",
            command=self.deselect_all,
            width=20
        )
        deselect_all_btn.pack(side=tk.LEFT)

        # Кнопка Применить - вторая строка
        buttons_frame_bottom = ttk.Frame(self.folders_frame)
        buttons_frame_bottom.pack(pady=(0, 15))

        apply_btn = ttk.Button(
            buttons_frame_bottom,
            text="Применить изменения",
            command=self.apply_folder_changes,
            width=20
        )
        apply_btn.pack()

    def check_initial_states(self):
        """Проверяет начальные состояния всех настроек"""
        self.check_bitlocker_state()
        self.check_photo_viewer_state()
        self.check_file_ext_state()
        self.check_recent_folders_state()

    def show_main_menu(self):
        """Показать главное меню"""
        self.folders_frame.pack_forget()
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.current_screen = "main"

    def show_folders_screen(self):
        """Показать экран управления папками"""
        self.main_canvas.pack_forget()
        self.main_scrollbar.pack_forget()
        self.folders_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.check_current_state()
        self.current_screen = "folders"

    def check_current_state(self):
        """Проверяет текущее состояние папок в реестре и обновляет индикаторы"""
        base_key = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"

        for guid in self.folders.keys():
            try:
                result = subprocess.run(
                    ['reg', 'query', f"{base_key}\\{guid}"],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                if result.returncode == 0:
                    self.checkbox_vars[guid].set(True)
                    self.status_labels[guid].config(
                        text="✓",
                        foreground='#27ae60'
                    )
                else:
                    self.checkbox_vars[guid].set(False)
                    self.status_labels[guid].config(
                        text="✗",
                        foreground='#e74c3c'
                    )

            except Exception:
                self.checkbox_vars[guid].set(False)
                self.status_labels[guid].config(
                    text="✗",
                    foreground='#e74c3c'
                )

    def check_bitlocker_state(self):
        """Проверяет текущее состояние BitLocker в реестре"""
        try:
            result1 = subprocess.run(
                ['reg', 'query', r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde', '/v', 'ProgrammaticAccessOnly'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            result2 = subprocess.run(
                ['reg', 'query', r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde-elev', '/v', 'ProgrammaticAccessOnly'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result1.returncode == 0 and result2.returncode == 0:
                self.bitlocker_enabled = False
                self.bitlocker_status.config(
                    text="Состояние: отключен (не отображается в меню)",
                    foreground='#e74c3c'
                )
                self.bitlocker_btn.config(text="Включить BitLocker в меню")
            else:
                self.bitlocker_enabled = True
                self.bitlocker_status.config(
                    text="Состояние: включен (отображается в меню)",
                    foreground='#27ae60'
                )
                self.bitlocker_btn.config(text="Отключить BitLocker из меню")

        except Exception:
            self.bitlocker_status.config(
                text="Состояние: ошибка проверки",
                foreground='#e67e22'
            )
            self.bitlocker_btn.config(text="Ошибка проверки")

    def check_photo_viewer_state(self):
        """Проверяет текущее состояние Windows Photo Viewer - более точная проверка"""
        try:
            # Проверяем наличие ключевой ассоциации для jpg файлов
            result = subprocess.run(
                ['reg', 'query', r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '/v',
                 '.jpg'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result.returncode == 0:
                # Также проверяем наличие класса PhotoViewer.FileAssoc.Jpeg
                result2 = subprocess.run(
                    ['reg', 'query', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                if result2.returncode == 0:
                    self.photo_viewer_enabled = True
                    self.photo_viewer_status.config(
                        text="Состояние: включен (форматы: .jpg, .png, .gif и др.)",
                        foreground='#27ae60'
                    )
                    self.photo_viewer_btn.config(text="Отключить Windows Photo Viewer")
                else:
                    # Есть ассоциации, но нет классов - неполная установка
                    self.photo_viewer_enabled = False
                    self.photo_viewer_status.config(
                        text="Состояние: частично установлен (требуется полное включение)",
                        foreground='#e67e22'
                    )
                    self.photo_viewer_btn.config(text="Включить Windows Photo Viewer полностью")
            else:
                self.photo_viewer_enabled = False
                self.photo_viewer_status.config(
                    text="Состояние: отключен (используется приложение 'Фотографии')",
                    foreground='#e74c3c'
                )
                self.photo_viewer_btn.config(text="Включить Windows Photo Viewer")

        except Exception:
            self.photo_viewer_status.config(
                text="Состояние: ошибка проверки",
                foreground='#e67e22'
            )
            self.photo_viewer_btn.config(text="Ошибка проверки")

    def check_file_ext_state(self):
        """Проверяет текущее состояние показа расширений файлов"""
        try:
            result = subprocess.run(
                ['reg', 'query', r'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced', '/v',
                 'HideFileExt'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result.returncode == 0:
                # Парсим значение из вывода
                for line in result.stdout.split('\n'):
                    if 'HideFileExt' in line:
                        value = line.split('REG_DWORD')[1].strip()
                        if value == '0x0':  # 0 в шестнадцатеричном формате
                            self.file_ext_enabled = True
                            self.file_ext_status.config(
                                text="Состояние: включен (расширения отображаются)",
                                foreground='#27ae60'
                            )
                            self.file_ext_btn.config(text="Скрыть расширения файлов")
                        else:  # 0x1 или другое значение
                            self.file_ext_enabled = False
                            self.file_ext_status.config(
                                text="Состояние: отключен (расширения скрыты)",
                                foreground='#e74c3c'
                            )
                            self.file_ext_btn.config(text="Показать расширения файлов")
                        break
            else:
                # Ключ не найден, используем значение по умолчанию (обычно скрыты)
                self.file_ext_enabled = False
                self.file_ext_status.config(
                    text="Состояние: отключен (расширения скрыты, по умолчанию)",
                    foreground='#e74c3c'
                )
                self.file_ext_btn.config(text="Показать расширения файлов")

        except Exception:
            self.file_ext_status.config(
                text="Состояние: ошибка проверки",
                foreground='#e67e22'
            )
            self.file_ext_btn.config(text="Ошибка проверки")

    def check_recent_folders_state(self):
        """Проверяет текущее состояние показа недавних папок"""
        try:
            result = subprocess.run(
                ['reg', 'query', r'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer', '/v',
                 'NoRecentDocsHistory'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result.returncode == 0:
                # Парсим значение из вывода
                for line in result.stdout.split('\n'):
                    if 'NoRecentDocsHistory' in line:
                        value = line.split('REG_DWORD')[1].strip()
                        if value == '0x1':  # 1 в шестнадцатеричном формате
                            self.recent_folders_enabled = False
                            self.recent_folders_status.config(
                                text="Состояние: отключен (недавние папки скрыты)",
                                foreground='#e74c3c'
                            )
                            self.recent_folders_btn.config(text="Включить недавние папки")
                        else:  # 0x0 или другое значение
                            self.recent_folders_enabled = True
                            self.recent_folders_status.config(
                                text="Состояние: включен (недавние папки отображаются)",
                                foreground='#27ae60'
                            )
                            self.recent_folders_btn.config(text="Отключить недавние папки")
                        break
            else:
                # Ключ не найден, значит недавние папки включены
                self.recent_folders_enabled = True
                self.recent_folders_status.config(
                    text="Состояние: включен (параметр отсутствует, по умолчанию)",
                    foreground='#27ae60'
                )
                self.recent_folders_btn.config(text="Отключить недавние папки")

        except Exception:
            self.recent_folders_status.config(
                text="Состояние: ошибка проверки",
                foreground='#e67e22'
            )
            self.recent_folders_btn.config(text="Ошибка проверки")

    def apply_folder_changes(self):
        """Применяет выбранные изменения для папок в реестре"""
        base_key = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"

        for guid, var in self.checkbox_vars.items():
            desired_state = var.get()
            reg_path = f"{base_key}\\{guid}"

            try:
                if desired_state:
                    subprocess.run(
                        ['reg', 'add', reg_path, '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    self.status_labels[guid].config(
                        text="✓",
                        foreground='#27ae60'
                    )
                else:
                    subprocess.run(
                        ['reg', 'delete', reg_path, '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    self.status_labels[guid].config(
                        text="✗",
                        foreground='#e74c3c'
                    )

            except Exception:
                if desired_state:
                    self.status_labels[guid].config(
                        text="✓",
                        foreground='#27ae60'
                    )
                else:
                    self.status_labels[guid].config(
                        text="✗",
                        foreground='#e74c3c'
                    )

    def toggle_bitlocker(self):
        """Переключает состояние BitLocker"""
        try:
            if self.bitlocker_enabled:
                action = "disable"
            else:
                action = "enable"

            keys = [
                r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde',
                r'HKEY_CLASSES_ROOT\Drive\shell\encrypt-bde-elev'
            ]

            for key in keys:
                try:
                    if action == "disable":
                        subprocess.run(
                            ['reg', 'add', key, '/v', 'ProgrammaticAccessOnly', '/t', 'REG_SZ', '/d', '', '/f'],
                            capture_output=True,
                            text=True,
                            encoding='cp866',
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    else:
                        subprocess.run(
                            ['reg', 'delete', key, '/v', 'ProgrammaticAccessOnly', '/f'],
                            capture_output=True,
                            text=True,
                            encoding='cp866',
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                except Exception:
                    pass

            self.bitlocker_enabled = not self.bitlocker_enabled
            self.check_bitlocker_state()

        except Exception:
            pass

    def toggle_photo_viewer(self):
        """Переключает состояние Windows Photo Viewer"""
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
        """Переключает состояние показа расширений файлов"""
        try:
            if self.file_ext_enabled:
                # Скрываем расширения (устанавливаем HideFileExt = 1)
                value = "1"
                action_text = "скрыть"
            else:
                # Показываем расширения (устанавливаем HideFileExt = 0)
                value = "0"
                action_text = "показать"

            reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

            subprocess.run(
                ['reg', 'add', reg_path, '/v', 'HideFileExt', '/t', 'REG_DWORD', '/d', value, '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            self.file_ext_enabled = not self.file_ext_enabled
            self.check_file_ext_state()

            print(f"Показ расширений файлов изменен: {action_text}")

        except Exception as e:
            print(f"Ошибка при переключении показа расширений: {e}")

    def toggle_recent_folders(self):
        """Переключает состояние показа недавних папок"""
        try:
            reg_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

            if self.recent_folders_enabled:
                # Отключаем недавние папки (устанавливаем NoRecentDocsHistory = 1)
                value = "1"
                action_text = "отключить"

                subprocess.run(
                    ['reg', 'add', reg_path, '/v', 'NoRecentDocsHistory', '/t', 'REG_DWORD', '/d', value, '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                # Включаем недавние папки (удаляем параметр NoRecentDocsHistory)
                action_text = "включить"

                subprocess.run(
                    ['reg', 'delete', reg_path, '/v', 'NoRecentDocsHistory', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

            self.recent_folders_enabled = not self.recent_folders_enabled
            self.check_recent_folders_state()

            print(f"Показ недавних папок изменен: {action_text}")

        except Exception as e:
            print(f"Ошибка при переключении показа недавних папок: {e}")

    def enable_photo_viewer(self):
        """Включает Windows Photo Viewer с полной конфигурацией"""
        try:
            # 1. Создаем ассоциации файлов
            reg_keys = [
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.jpg',
                 'PhotoViewer.FileAssoc.Jpeg'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.wdp',
                 'PhotoViewer.FileAssoc.Wdp'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.jfif',
                 'PhotoViewer.FileAssoc.JFIF'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.dib',
                 'PhotoViewer.FileAssoc.Bitmap'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.png',
                 'PhotoViewer.FileAssoc.Png'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.tiff',
                 'PhotoViewer.FileAssoc.Tiff'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.jxr',
                 'PhotoViewer.FileAssoc.Wdp'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.bmp',
                 'PhotoViewer.FileAssoc.Bitmap'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.jpe',
                 'PhotoViewer.FileAssoc.Jpeg'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.jpeg',
                 'PhotoViewer.FileAssoc.Jpeg'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.gif',
                 'PhotoViewer.FileAssoc.Gif'],
                [r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.tif',
                 'PhotoViewer.FileAssoc.Tiff'],
            ]

            for key, ext, value in reg_keys:
                subprocess.run(
                    ['reg', 'add', key, '/v', ext, '/t', 'REG_SZ', '/d', value, '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

            # 2. Создаем основные классы PhotoViewer.FileAssoc.*
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
                subprocess.run(
                    ['reg', 'add', class_key, '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

            # 3. Добавляем дополнительные параметры для каждого класса
            # Для PhotoViewer.FileAssoc.Bitmap
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap', '/v', 'ImageOptionFlags', '/t',
                 'REG_DWORD', '/d', '1', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # Для PhotoViewer.FileAssoc.JFIF
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF', '/v', 'EditFlags', '/t',
                 'REG_DWORD', '/d', '65536', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF', '/v', 'ImageOptionFlags', '/t',
                 'REG_DWORD', '/d', '1', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # Для PhotoViewer.FileAssoc.Jpeg
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg', '/v', 'EditFlags', '/t',
                 'REG_DWORD', '/d', '65536', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg', '/v', 'ImageOptionFlags', '/t',
                 'REG_DWORD', '/d', '1', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # Для PhotoViewer.FileAssoc.Png
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png', '/v', 'ImageOptionFlags', '/t',
                 'REG_DWORD', '/d', '1', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # Для PhotoViewer.FileAssoc.Wdp
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp', '/v', 'EditFlags', '/t', 'REG_DWORD',
                 '/d', '65536', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp', '/v', 'ImageOptionFlags', '/t',
                 'REG_DWORD', '/d', '1', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # 4. Создаем структуру shell\open для каждого класса
            classes_with_shell = [
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Gif',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Tiff'
            ]

            for class_key in classes_with_shell:
                # Создаем структуру shell
                subprocess.run(
                    ['reg', 'add', class_key + r'\shell', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                # Создаем структуру shell\open
                subprocess.run(
                    ['reg', 'add', class_key + r'\shell\open', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                # Создаем структуру shell\open\command
                subprocess.run(
                    ['reg', 'add', class_key + r'\shell\open\command', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                # Устанавливаем команду открытия
                command_value = r'%SystemRoot%\System32\rundll32.exe "%ProgramFiles%\Windows Photo Viewer\PhotoViewer.dll", ImageView_Fullscreen %1'
                subprocess.run(
                    ['reg', 'add', class_key + r'\shell\open\command', '/ve', '/t', 'REG_SZ', '/d', command_value,
                     '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

                # Добавляем DropTarget
                subprocess.run(
                    ['reg', 'add', class_key + r'\shell\open\DropTarget', '/v', 'Clsid', '/t', 'REG_SZ', '/d',
                     '{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

            # 5. Добавляем DefaultIcon для некоторых классов
            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap\DefaultIcon', '/ve', '/t', 'REG_SZ',
                 '/d', '%SystemRoot%\\System32\\imageres.dll,-70', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF\DefaultIcon', '/ve', '/t', 'REG_SZ',
                 '/d', '%SystemRoot%\\System32\\imageres.dll,-72', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg\DefaultIcon', '/ve', '/t', 'REG_SZ',
                 '/d', '%SystemRoot%\\System32\\imageres.dll,-72', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png\DefaultIcon', '/ve', '/t', 'REG_SZ',
                 '/d', '%SystemRoot%\\System32\\imageres.dll,-71', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            subprocess.run(
                ['reg', 'add', r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp\DefaultIcon', '/ve', '/t', 'REG_SZ',
                 '/d', '%SystemRoot%\\System32\\wmphoto.dll,-400', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            print("Windows Photo Viewer включен со всеми необходимыми параметрами")

        except Exception as e:
            print(f"Ошибка при включении Photo Viewer: {e}")

    def disable_photo_viewer(self):
        """Отключает Windows Photo Viewer"""
        try:
            # 1. Удаляем ассоциации файлов
            extensions = ['.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff', '.wdp', '.jfif', '.jpe', '.dib',
                          '.jxr']

            for ext in extensions:
                try:
                    subprocess.run(
                        ['reg', 'delete', r'HKLM\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations',
                         '/v', ext, '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                except:
                    pass

            # 2. Удаляем классы PhotoViewer.FileAssoc.*
            photo_viewer_classes = [
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Bitmap',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.JFIF',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Jpeg',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Png',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Wdp',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Gif',
                r'HKLM\SOFTWARE\Classes\PhotoViewer.FileAssoc.Tiff',
                r'HKLM\SOFTWARE\Classes\Applications\photoviewer.dll'
            ]

            for class_key in photo_viewer_classes:
                try:
                    subprocess.run(
                        ['reg', 'delete', class_key, '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                except:
                    pass

            # 3. Также удаляем из HKCU для надежности
            for ext in ['.bmp', '.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff', '.wdp', '.jfif']:
                try:
                    subprocess.run(
                        ['reg', 'delete', fr'HKCU\SOFTWARE\Classes\{ext}', '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                except:
                    pass

            print("Windows Photo Viewer отключен")

        except Exception as e:
            print(f"Ошибка при отключении Photo Viewer: {e}")

    def select_all(self):
        """Выбрать все папки"""
        for guid, var in self.checkbox_vars.items():
            var.set(True)

    def deselect_all(self):
        """Снять выделение со всех папок"""
        for guid, var in self.checkbox_vars.items():
            var.set(False)

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()