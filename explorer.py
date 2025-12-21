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

        # Для хранения индикаторов состояния
        self.status_labels = {}

        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self.create_widgets()
        self.check_current_state()
        self.check_bitlocker_state()
        self.show()

    def create_widgets(self):
        """Создание виджетов для управления папками в 'Компьютер'"""
        title_label = ttk.Label(
            self.content_frame,
            text="Настройки Проводника",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        notebook = ttk.Notebook(self.content_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)

        folders_tab = ttk.Frame(notebook)
        notebook.add(folders_tab, text="        Папки в Этот компьютер        ")

        bitlocker_tab = ttk.Frame(notebook)
        notebook.add(bitlocker_tab, text="        Контекстное меню дисков        ")

        self.create_folders_tab(folders_tab)
        self.create_bitlocker_tab(bitlocker_tab)

    def create_folders_tab(self, parent):
        """Создание вкладки управления папками"""
        desc_label = ttk.Label(
            parent,
            text="Выберите какие папки отображать\n в разделе Этот компьютер",
            font=('Arial', 10),
            foreground='#7f8c8d',
            wraplength=500
        )
        desc_label.pack(pady=(0, 15))

        folders_frame = ttk.LabelFrame(parent, text="Доступные папки", padding=20)
        folders_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=10)

        canvas = tk.Canvas(folders_frame, highlightthickness=0, height=200)
        scrollbar = ttk.Scrollbar(folders_frame, orient="vertical", command=canvas.yview)
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

            # Создаем чекбокс
            checkbox = ttk.Checkbutton(
                frame,
                text=name,
                variable=var,
                onvalue=True,
                offvalue=False
            )
            checkbox.pack(side=tk.LEFT, anchor=tk.W)

            # Создаем индикатор состояния
            status_label = ttk.Label(
                frame,
                text="•",
                font=('Arial', 9, 'bold'),
                width=2
            )
            status_label.pack(side=tk.LEFT, padx=(2, 0))
            self.status_labels[guid] = status_label

            # Описание папки
            desc_label = ttk.Label(
                frame,
                text=f" - {description}",
                font=('Arial', 9),
                foreground='#95a5a6'
            )
            desc_label.pack(side=tk.LEFT, padx=(0, 8))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        folders_buttons_frame = ttk.Frame(parent)
        folders_buttons_frame.pack(pady=15)

        center_frame = ttk.Frame(folders_buttons_frame)
        center_frame.pack(expand=True)

        select_all_btn = ttk.Button(
            center_frame,
            text="Выбрать все папки",
            command=self.select_all,
            width=20
        )
        select_all_btn.pack(side=tk.LEFT, padx=(0, 15))

        deselect_all_btn = ttk.Button(
            center_frame,
            text="Снять все папки",
            command=self.deselect_all,
            width=20
        )
        deselect_all_btn.pack(side=tk.LEFT)

        apply_frame = ttk.Frame(parent)
        apply_frame.pack(pady=(10, 0))

        center_apply_frame = ttk.Frame(apply_frame)
        center_apply_frame.pack(expand=True)

        apply_folders_btn = ttk.Button(
            center_apply_frame,
            text="Применить для папок",
            command=self.apply_folder_changes,
            width=25
        )
        apply_folders_btn.pack()

    def create_bitlocker_tab(self, parent):
        """Создание вкладки управления BitLocker"""
        desc_label = ttk.Label(
            parent,
            text="Управление пунктом BitLocker в контекстном меню дисков",
            font=('Arial', 10),
            foreground='#7f8c8d',
            wraplength=500
        )
        desc_label.pack(pady=(0, 20))

        bitlocker_frame = ttk.LabelFrame(parent, text="Настройки BitLocker", padding=25)
        bitlocker_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=10)

        status_container = ttk.Frame(bitlocker_frame)
        status_container.pack(expand=True, fill=tk.BOTH, pady=(0, 25))

        status_center_frame = ttk.Frame(status_container)
        status_center_frame.place(relx=0.5, rely=0.5, anchor='center')

        self.bitlocker_status_label = ttk.Label(
            status_center_frame,
            text="Состояние: проверяется...",
            font=('Arial', 10),
            foreground='#95a5a6'
        )
        self.bitlocker_status_label.pack()

        button_frame = ttk.Frame(bitlocker_frame)
        button_frame.pack(expand=True, pady=10)

        center_button_frame = ttk.Frame(button_frame)
        center_button_frame.pack(expand=True)

        self.bitlocker_button = ttk.Button(
            center_button_frame,
            text="---",
            command=self.toggle_bitlocker,
            width=25
        )
        self.bitlocker_button.pack()

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
                    # Обновляем индикатор - галочка (включено)
                    self.status_labels[guid].config(
                        text="✓",
                        foreground='#27ae60'  # Зеленый цвет
                    )
                else:
                    self.checkbox_vars[guid].set(False)
                    # Обновляем индикатор - крестик (выключено)
                    self.status_labels[guid].config(
                        text="✗",
                        foreground='#e74c3c'  # Красный цвет
                    )

            except Exception:
                self.checkbox_vars[guid].set(False)
                # В случае ошибки - крестик
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
                self.bitlocker_status_label.config(
                    text="BitLocker СКРЫТ из контекстного меню",
                    foreground='#e74c3c'
                )
                self.bitlocker_button.config(text="Включить BitLocker")
            else:
                self.bitlocker_enabled = True
                self.bitlocker_status_label.config(
                    text="BitLocker ОТОБРАЖАЕТСЯ в контекстном меню",
                    foreground='#27ae60'
                )
                self.bitlocker_button.config(text="Отключить BitLocker")

        except Exception:
            self.bitlocker_status_label.config(
                text="Состояние: ошибка проверки",
                foreground='#e67e22'
            )
            self.bitlocker_button.config(text="Ошибка проверки")

    def apply_folder_changes(self):
        """Применяет выбранные изменения для папок в реестре и сразу обновляет индикаторы"""
        base_key = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace"

        # Применяем изменения и сразу обновляем индикаторы
        for guid, var in self.checkbox_vars.items():
            desired_state = var.get()
            reg_path = f"{base_key}\\{guid}"

            try:
                # Применяем изменения в реестре
                if desired_state:
                    # Включаем папку - добавляем ключ в реестр
                    subprocess.run(
                        ['reg', 'add', reg_path, '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    # Сразу обновляем индикатор на галочку (зеленый)
                    self.status_labels[guid].config(
                        text="✓",
                        foreground='#27ae60'
                    )
                else:
                    # Выключаем папку - удаляем ключ из реестра
                    subprocess.run(
                        ['reg', 'delete', reg_path, '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    # Сразу обновляем индикатор на крестик (красный)
                    self.status_labels[guid].config(
                        text="✗",
                        foreground='#e74c3c'
                    )

            except Exception:
                # В случае ошибки при применении, показываем текущее состояние чекбокса
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

    def select_all(self):
        """Выбрать все папки"""
        for guid, var in self.checkbox_vars.items():
            var.set(True)
            # Не меняем индикатор, только чекбокс

    def deselect_all(self):
        """Снять выделение со всех папок"""
        for guid, var in self.checkbox_vars.items():
            var.set(False)
            # Не меняем индикатор, только чекбокс

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()