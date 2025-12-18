import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
import os
import subprocess

# Вместо from tkinter import filedialog, используем:
import tkinter.filedialog as filedialog


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)

        # Переменная для хранения пути к выбранной папке
        self.selected_folder = None

        # Словарь для хранения состояний чекбоксов (все выключены по умолчанию)
        self.file_vars = {
            "imageres.dll": tk.BooleanVar(value=False),
            "imagesp1.dll": tk.BooleanVar(value=False),
            "networkexplorer.dll": tk.BooleanVar(value=False),
            "DDORes.dll": tk.BooleanVar(value=False),
            "mmres.dll": tk.BooleanVar(value=False)
        }

        # Эталонные хеши для всех файлов
        self.expected_hashes = {
            "imageres.dll": "DDD5D7462A26F7837CD02A7667615354F9377A0A727BF880D825676D6DDC4CB7",
            "imagesp1.dll": "466E64D491C7B3D7051FBDC0907CD813D7A6435D18BF4B00CE0C3191B80E0541",
            "networkexplorer.dll": "7BF68211CFBDDE2E377C992CEC2F5A0C1F41B1F12BE3A99733A549AB63A8C961",
            "DDORes.dll": "4802C8F0BAAE3765FDFF16F09F56647B1C8B20889E4C531B773442EC068CBC3B",
            "mmres.dll": "4651BB26196D0ED493A26F4E401C8FAC0465E7F9DE2ECFED1642155A6EAF3703"
        }

        # Ссылки на кнопки
        self.start_button = None
        self.restore_button = None

        # Словарь для хранения статусов файлов в системе
        self.system_file_status = {}

        # Сначала создаем виджеты
        self.create_widgets()

        # Затем проверяем статусы и обновляем кнопки
        self.check_system_files_status()
        self.update_buttons_state()

        # Показываем модуль
        self.show()

    def create_tooltip(self, widget, text):
        """Создает всплывающую подсказку для виджета."""

        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")

            label = tk.Label(tooltip, text=text, background="lightyellow",
                             relief="solid", borderwidth=1, font=("Arial", 8))
            label.pack()
            widget.tooltip = tooltip

        def on_leave(event):
            if hasattr(widget, 'tooltip') and widget.tooltip:
                widget.tooltip.destroy()
                widget.tooltip = None

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def create_widgets(self):
        """Создание виджетов для модификации системных DLL-файлов"""
        # Заголовок
        title_label = ttk.Label(
            self.frame,
            text="Модификация системных DLL-файлов",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        # Описание
        desc_label = ttk.Label(
            self.frame,
            text="Данная функция позволяет добавить в систему модифицированные файлы\n"
                 "содержащие иконки от Windows 7. Выберите файлы для модификации:",
            font=('Arial', 10),
            foreground='#7f8c8d',
            justify="center"
        )
        desc_label.pack(pady=(0, 20))

        # Фрейм для настроек
        settings_frame = ttk.LabelFrame(self.frame, text="Настройки модификации", padding=15)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Кнопка "Выбрать папку с модифицированными файлами"
        self.select_folder_button = ttk.Button(
            settings_frame,
            text="Выбрать папку с модифицированными файлами",
            command=self.select_folder,
            width=40
        )
        self.select_folder_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.select_folder_button, "Выберите папку, содержащую модифицированные DLL-файлы")

        # Метка для отображения результата проверки
        self.result_label = ttk.Label(settings_frame, text="", foreground="red")
        self.result_label.pack(pady=5)

        # Фрейм для чекбоксов
        checkbox_frame = ttk.Frame(settings_frame)
        checkbox_frame.pack(pady=10, fill=tk.X)

        # Заголовок для чекбоксов
        ttk.Label(checkbox_frame, text="Выберите файлы для модификации:",
                  font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))

        # Чекбоксы для каждого файла
        files_info = {
            "imageres.dll": "Основные иконки системы",
            "imagesp1.dll": "Дополнительные иконки",
            "networkexplorer.dll": "Иконки сети и проводника",
            "DDORes.dll": "Иконки DDE",
            "mmres.dll": "Иконки звука"
        }

        for filename, description in files_info.items():
            file_frame = ttk.Frame(checkbox_frame)
            file_frame.pack(fill=tk.X, pady=2)

            checkbox = ttk.Checkbutton(
                file_frame,
                text=f"{filename} - {description}",
                variable=self.file_vars[filename],
                command=self.update_buttons_state
            )
            checkbox.pack(side=tk.LEFT, anchor=tk.W)
            self.create_tooltip(checkbox, f"Включить/выключить модификацию файла {filename}")

            # Метка для отображения статуса файла в системе
            system_status_label = ttk.Label(file_frame, text="", font=('Arial', 8))
            system_status_label.pack(side=tk.RIGHT, padx=5)
            # Сохраняем ссылку на метку статуса системы
            setattr(self, f"{filename.replace('.', '_')}_system_status", system_status_label)

        # Фрейм для кнопок действий
        action_frame = ttk.Frame(settings_frame)
        action_frame.pack(pady=15, fill=tk.X)

        # Кнопка "Начать"
        self.start_button = ttk.Button(
            action_frame,
            text="Начать",
            command=self.start_modification,
            state="disabled"
        )
        self.start_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.create_tooltip(self.start_button, "Начать процесс модификации выбранных файлов")

        # Кнопка "Восстановить"
        self.restore_button = ttk.Button(
            action_frame,
            text="Восстановить",
            command=self.restore_originals,
            state="disabled"
        )
        self.restore_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.create_tooltip(self.restore_button, "Восстановить оригинальные версии файлов")

        # Кнопка "Посмотреть хеш"
        view_hash_button = ttk.Button(
            action_frame,
            text="Посмотреть хеш",
            command=self.view_hashes
        )
        view_hash_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.create_tooltip(view_hash_button, "Показать хеш-суммы файлов в системе")

        # Информационная панель
        info_label = ttk.Label(
            settings_frame,
            text="⚠️ Внимание: Модификация системных файлов \nвыполняется на свой страх и риск! \nНастоятельно рекомендуется сделать точку восстановления системы",
            font=('Arial', 8),
            foreground='#e67e22',
            justify=tk.CENTER
        )
        info_label.pack(pady=10)

    def check_system_files_status(self):
        """Проверяет статус файлов в системе и обновляет отображение."""
        self.system_file_status = {}

        for filename, expected_hash in self.expected_hashes.items():
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            status_label = getattr(self, f"{filename.replace('.', '_')}_system_status")

            try:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as file:
                        file_content = file.read()
                        file_hash = hashlib.sha256(file_content).hexdigest().lower()

                    is_match = file_hash == expected_hash.lower()

                    # Сохраняем статус файла в системе
                    self.system_file_status[filename] = {
                        'exists': True,
                        'matches_expected': is_match,
                        'hash': file_hash
                    }

                    # Обновляем отображение статуса
                    if is_match:
                        status_label.config(text="✓ В системе", foreground="green")
                    else:
                        status_label.config(text="✗ Оригинал", foreground="red")
                else:
                    self.system_file_status[filename] = {
                        'exists': False,
                        'matches_expected': False,
                        'hash': None
                    }
                    status_label.config(text="✗ Отсутствует", foreground="red")
            except Exception as e:
                self.system_file_status[filename] = {
                    'exists': False,
                    'matches_expected': False,
                    'hash': None,
                    'error': str(e)
                }
                status_label.config(text="✗ Ошибка", foreground="red")

    def update_buttons_state(self):
        """Обновляет состояние кнопок в зависимости от выбранных чекбоксов и статуса файлов."""
        # Получаем список выбранных файлов
        selected_files = [filename for filename, var in self.file_vars.items() if var.get()]

        if not selected_files:
            # Если ничего не выбрано - обе кнопки отключены
            self.start_button.config(state="disabled")
            self.restore_button.config(state="disabled")
            return

        # Проверяем статусы выбранных файлов
        any_selected_mismatch = False
        any_selected_match = False
        all_selected_exist = True

        for filename in selected_files:
            file_status = self.system_file_status.get(filename, {})
            exists = file_status.get('exists', False)
            matches = file_status.get('matches_expected', False)

            if not exists:
                all_selected_exist = False

            if matches:
                any_selected_match = True
            else:
                any_selected_mismatch = True

        # Логика блокировки кнопок:
        # - "Начать" доступна только если есть несовпадающие файлы И папка выбрана
        # - "Восстановить" доступна только если есть совпадающие файлы
        start_enabled = any_selected_mismatch and all_selected_exist and self.selected_folder
        restore_enabled = any_selected_match and all_selected_exist

        self.start_button.config(state="normal" if start_enabled else "disabled")
        self.restore_button.config(state="normal" if restore_enabled else "disabled")

    def select_folder(self):
        """Открывает диалог выбора папки с модифицированными файлами."""
        folder_path = filedialog.askdirectory()
        if folder_path:
            # Проверка на запрещенную папку System32
            system32_path = os.path.join(os.environ["WINDIR"], "System32")
            if os.path.normcase(folder_path) == os.path.normcase(system32_path):
                messagebox.showerror("Ошибка", "Выбор папки System32 запрещен!")
                return

            self.selected_folder = folder_path
            self.log_to_file(f"Выбрана папка: {folder_path}")
            # Автоматически проверяем файлы после выбора папки
            self.check_files_and_show_result()
            # Обновляем состояние кнопок
            self.update_buttons_state()

    def check_files_and_show_result(self):
        """Проверяет файлы в выбранной папке на соответствие хешам и показывает результат."""
        if not self.selected_folder:
            self.result_label.config(text="Папка не выбрана", foreground="red")
            return

        all_files_exist = True
        all_files_match = True

        # Проверяем ВСЕ файлы независимо от состояния чекбоксов
        for filename, expected_hash in self.expected_hashes.items():
            file_path = os.path.join(self.selected_folder, filename)

            if os.path.exists(file_path):
                with open(file_path, "rb") as file:
                    file_content = file.read()
                    file_hash = hashlib.sha256(file_content).hexdigest().lower()

                if file_hash == expected_hash.lower():
                    self.log_to_file(f"Файл {filename} соответствует эталонному хешу.")
                else:
                    all_files_match = False
                    self.log_to_file(f"Файл {filename} НЕ соответствует эталонному хешу.")
            else:
                all_files_exist = False
                self.log_to_file(f"Файл {filename} не найден.")

        # Общий результат
        if all_files_exist and all_files_match:
            self.result_label.config(text="✓ Все файлы присутствуют и соответствуют эталонным хешам",
                                     foreground="green")
        elif all_files_exist and not all_files_match:
            self.result_label.config(text="✓ Все файлы присутствуют, но некоторые не соответствуют хешам",
                                     foreground="orange")
        else:
            self.result_label.config(text="✗ Некоторые файлы отсутствуют или не соответствуют", foreground="red")

    def log_to_file(self, message):
        """Записывает сообщение в файл mod.log на рабочем столе."""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            log_file_path = os.path.join(desktop_path, "mod.log")

            # Создаем файл если он не существует
            os.makedirs(os.path.dirname(log_file_path), exist_ok=True)

            with open(log_file_path, "a", encoding="utf-8") as log_file:
                log_file.write(message + "\n")
        except Exception as e:
            print(f"Ошибка записи в лог: {e}")

    def restart_explorer(self):
        """Запускает explorer.exe, если он не запущен."""
        try:
            # Проверяем, запущен ли explorer.exe
            result = subprocess.run('tasklist', shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            if "explorer.exe" not in result.stdout.lower():
                self.log_to_file("Запуск explorer.exe...")
                subprocess.Popen("explorer.exe", shell=True)
            else:
                self.log_to_file("Explorer.exe уже запущен.")
        except Exception as e:
            self.log_to_file(f"Ошибка при запуске explorer.exe: {e}")

    def restore_originals(self):
        """Восстанавливает оригинальные системные DLL-файлы."""
        # Получаем список файлов для восстановления на основе выбранных чекбоксов
        files_to_process = [filename for filename, var in self.file_vars.items() if var.get()]

        if not files_to_process:
            messagebox.showwarning("Предупреждение", "Не выбраны файлы для восстановления.")
            return

        # Проверяем, что выбранные файлы действительно нуждаются в восстановлении
        files_need_restore = []
        for filename in files_to_process:
            file_status = self.system_file_status.get(filename, {})
            if file_status.get('exists') and file_status.get('matches_expected'):
                files_need_restore.append(filename)

        if not files_need_restore:
            messagebox.showinfo("Информация", "Выбранные файлы уже являются оригинальными и не требуют восстановления.")
            return

        # Завершение explorer.exe
        try:
            self.log_to_file("Завершение процесса explorer.exe...")
            subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при завершении explorer.exe: {e}")

        # Восстановление файлов
        restored_count = 0
        for filename in files_need_restore:
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            if os.path.exists(file_path):
                self.log_to_file(f"Попытка удаления файла {filename}...")
                try:
                    # Попытка удаления файла
                    subprocess.run(f'del "{file_path}" >nul 2>&1', shell=True, check=True)
                    self.log_to_file(f"Файл {filename} успешно удален.")
                except Exception as e:
                    self.log_to_file(f"Не удалось удалить файл {filename}. Переименовываем его...")
                    # Переименование в .dll_mod
                    mod_file_path = f"{file_path}_mod"
                    try:
                        os.rename(file_path, mod_file_path)
                        self.log_to_file(f"Файл {filename} переименован в {os.path.basename(mod_file_path)}.")
                        # Попытка повторного удаления
                        try:
                            subprocess.run(f'del "{mod_file_path}" >nul 2>&1', shell=True, check=True)
                            self.log_to_file(f"Файл {os.path.basename(mod_file_path)} успешно удален.")
                        except Exception as e:
                            self.log_to_file(
                                f"Не удалось удалить временный файл {os.path.basename(mod_file_path)}. Пропускаем...")
                    except Exception as e:
                        self.log_to_file(f"Ошибка при переименовании файла {filename}: {e}")

            # Восстановление оригинального файла
            orig_file_path = f"{file_path}_orig"
            if os.path.exists(orig_file_path):
                try:
                    os.rename(orig_file_path, file_path)
                    self.log_to_file(f"Оригинальный файл {filename} успешно восстановлен.")
                    restored_count += 1
                except Exception as e:
                    self.log_to_file(f"Ошибка при восстановлении файла {filename}: {e}")
                    # Если файл уже существует, переименовываем его в .dll_mod и пытаемся снова
                    if os.path.exists(file_path):
                        existing_file_path = f"{file_path}_mod"
                        try:
                            os.rename(file_path, existing_file_path)
                            self.log_to_file(f"Файл {filename} переименован в {os.path.basename(existing_file_path)}.")
                            # Повторная попытка восстановления
                            try:
                                os.rename(orig_file_path, file_path)
                                self.log_to_file(
                                    f"Оригинальный файл {filename} успешно восстановлен после переименования.")
                                restored_count += 1
                            except Exception as e:
                                self.log_to_file(f"Ошибка при повторной попытке восстановления файла {filename}: {e}")
                        except Exception as e:
                            self.log_to_file(f"Ошибка при переименовании существующего файла {filename}: {e}")

            # Восстановление прав доступа
            acl_file = f"{file_path}.acl"
            if os.path.exists(acl_file):
                try:
                    subprocess.run(f'icacls "{file_path}" /restore "{acl_file}" >nul 2>&1', shell=True, check=True)
                    self.log_to_file(f"Права доступа для {filename} успешно восстановлены.")
                except Exception as e:
                    self.log_to_file(f"Ошибка при восстановлении прав доступа для {filename}: {e}")

        # Запуск explorer.exe
        try:
            self.log_to_file("Запуск explorer.exe...")
            subprocess.Popen("explorer.exe", shell=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при запуске explorer.exe: {e}")

        # Обновляем статус файлов после восстановления
        self.check_system_files_status()
        self.update_buttons_state()

        # Уведомление пользователя
        messagebox.showinfo("Успешно",
                            f"Восстановление оригинальных DLL-файлов завершено успешно.\nВосстановлено файлов: {restored_count}")

    def start_modification(self):
        """Запускает процесс модификации."""
        if not self.selected_folder:
            messagebox.showerror("Ошибка", "Выберите папку с модифицированными файлами.")
            return

        # Получаем список файлов для модификации на основе выбранных чекбоксов
        files_to_process = [filename for filename, var in self.file_vars.items() if var.get()]

        if not files_to_process:
            messagebox.showwarning("Предупреждение", "Не выбраны файлы для модификации.")
            return

        # Проверяем, что выбранные файлы действительно нуждаются в модификации
        files_need_modification = []
        for filename in files_to_process:
            file_status = self.system_file_status.get(filename, {})
            if file_status.get('exists') and not file_status.get('matches_expected'):
                files_need_modification.append(filename)

        if not files_need_modification:
            messagebox.showinfo("Информация",
                                "Выбранные файлы уже соответствуют эталонным хешам и не требуют модификации.")
            return

        # Проверка наличия всех необходимых файлов в выбранной папке
        missing_files = []
        for filename in files_need_modification:
            source_file = os.path.join(self.selected_folder, filename)
            if not os.path.exists(source_file):
                missing_files.append(filename)
        if missing_files:
            messagebox.showerror("Ошибка",
                                 f"Файлы отсутствуют в выбранной папке: {', '.join(missing_files)}!")
            return

        # Завершение explorer.exe
        try:
            self.log_to_file("Завершение процесса explorer.exe...")
            subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при завершении explorer.exe: {e}")

        # Основной процесс модификации
        modified_count = 0
        for filename in files_need_modification:
            source_file = os.path.join(self.selected_folder, filename)
            target_file = os.path.join(os.environ["WINDIR"], "System32", filename)

            try:
                # Сохраняем права доступа
                acl_file = f"{target_file}.acl"
                subprocess.run(f'icacls "{target_file}" /save "{acl_file}" >nul 2>&1', shell=True, check=True)
                # Получаем полный доступ к файлу
                subprocess.run(f'takeown /f "{target_file}" /a >nul 2>&1', shell=True, check=True)
                subprocess.run(f'icacls "{target_file}" /grant "%USERNAME%":(F) >nul 2>&1', shell=True, check=True)

                # Переименовываем оригинальный файл
                backup_file = f"{target_file}_orig"
                if os.path.exists(target_file):
                    os.rename(target_file, backup_file)
                    self.log_to_file(f"Оригинальный файл {filename} переименован в {os.path.basename(backup_file)}.")

                # Копируем новый файл
                subprocess.run(f'copy "{source_file}" "{target_file}" >nul 2>&1', shell=True, check=True)
                self.log_to_file(f"Файл {filename} успешно заменён.")
                modified_count += 1

                # Восстанавливаем права доступа
                subprocess.run(f'icacls "{target_file}" /setowner "NT SERVICE\\TrustedInstaller" >nul 2>&1', shell=True,
                               check=True)
                subprocess.run(f'icacls "{target_file}" /restore "{acl_file}" >nul 2>&1', shell=True, check=True)
            except Exception as e:
                self.log_to_file(f"Ошибка при обработке файла {filename}: {e}")

        # Запуск explorer.exe
        try:
            self.log_to_file("Запуск explorer.exe...")
            subprocess.Popen("explorer.exe", shell=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при запуске explorer.exe: {e}")

        # Обновляем статус файлов после модификации
        self.check_system_files_status()
        self.update_buttons_state()

        # Уведомление пользователя
        messagebox.showinfo("Успешно",
                            f"Модификация системных DLL файлов завершена успешно.\nОбработано файлов: {modified_count}")

    def view_hashes(self):
        """Показывает новое окно с хеш-суммами файлов."""
        # Создание нового окна
        hash_window = tk.Toplevel(self.parent)
        hash_window.title("Хеш-суммы файлов")
        hash_window.geometry("900x210")  # Увеличил ширину окна
        hash_window.resizable(False, False)

        # Главный фрейм
        main_frame = ttk.Frame(hash_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        header_label = ttk.Label(main_frame, text="Хеш-суммы файлов в системе:", font=("Arial", 14, "bold"))
        header_label.pack(anchor=tk.W, pady=(0, 10))

        # Фрейм для содержимого
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовки колонок
        header_frame = ttk.Frame(content_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(header_frame, text="Файл", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(header_frame, text="Статус", font=("Arial", 10, "bold"), width=20).pack(side=tk.LEFT, padx=(
        0, 10))  # Увеличил ширину
        ttk.Label(header_frame, text="SHA-256 Хеш", font=("Arial", 10, "bold")).pack(side=tk.LEFT)

        # Отображение хешей для каждого файла
        for filename, expected_hash in self.expected_hashes.items():
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)

            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            try:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as file:
                        file_content = file.read()
                        file_hash = hashlib.sha256(file_content).hexdigest().upper()
                    is_match = file_hash == expected_hash.upper()

                    status_text = "✓ Соответствует" if is_match else "✗ Не соответствует"
                    status_color = "green" if is_match else "red"

                    # Название файла
                    file_label = ttk.Label(row_frame, text=filename, font=("Courier New", 9), width=25, anchor="w")
                    file_label.pack(side=tk.LEFT, padx=(0, 10))

                    # Статус
                    status_label = ttk.Label(row_frame, text=status_text, foreground=status_color,
                                             font=("Arial", 9), width=20, anchor="w")  # Увеличил ширину
                    status_label.pack(side=tk.LEFT, padx=(0, 10))

                    # Хеш
                    hash_label = ttk.Label(row_frame, text=file_hash, font=("Courier New", 9), anchor="w")
                    hash_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
                else:
                    # Название файла
                    file_label = ttk.Label(row_frame, text=filename, font=("Courier New", 9), width=25, anchor="w")
                    file_label.pack(side=tk.LEFT, padx=(0, 10))

                    # Статус
                    status_label = ttk.Label(row_frame, text="✗ Не найден", foreground="red",
                                             font=("Arial", 9), width=20, anchor="w")  # Увеличил ширину
                    status_label.pack(side=tk.LEFT, padx=(0, 10))

                    # Сообщение об ошибке
                    hash_label = ttk.Label(row_frame, text="Файл не существует", font=("Courier New", 9),
                                           foreground="red", anchor="w")
                    hash_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
            except Exception as e:
                # Название файла
                file_label = ttk.Label(row_frame, text=filename, font=("Courier New", 9), width=25, anchor="w")
                file_label.pack(side=tk.LEFT, padx=(0, 10))

                # Статус
                status_label = ttk.Label(row_frame, text="✗ Ошибка", foreground="red",
                                         font=("Arial", 9), width=20, anchor="w")  # Увеличил ширину
                status_label.pack(side=tk.LEFT, padx=(0, 10))

                # Сообщение об ошибке
                error_label = ttk.Label(row_frame, text=f"Ошибка чтения: {str(e)}", font=("Courier New", 9),
                                        foreground="red", anchor="w")
                error_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()