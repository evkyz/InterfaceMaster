import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
import os
import subprocess
import tkinter.filedialog as filedialog


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)

        self.selected_folder = None

        self.file_vars = {
            "imageres.dll": tk.BooleanVar(value=False),
            "imagesp1.dll": tk.BooleanVar(value=False),
            "networkexplorer.dll": tk.BooleanVar(value=False),
            "DDORes.dll": tk.BooleanVar(value=False),
            "mmres.dll": tk.BooleanVar(value=False)
        }

        self.expected_hashes = {
            "imageres.dll": "DDD5D7462A26F7837CD02A7667615354F9377A0A727BF880D825676D6DDC4CB7",
            "imagesp1.dll": "466E64D491C7B3D7051FBDC0907CD813D7A6435D18BF4B00CE0C3191B80E0541",
            "networkexplorer.dll": "7BF68211CFBDDE2E377C992CEC2F5A0C1F41B1F12BE3A99733A549AB63A8C961",
            "DDORes.dll": "4802C8F0BAAE3765FDFF16F09F56647B1C8B20889E4C531B773442EC068CBC3B",
            "mmres.dll": "4651BB26196D0ED493A26F4E401C8FAC0465E7F9DE2ECFED1642155A6EAF3703"
        }

        self.start_button = None
        self.restore_button = None
        self.system_file_status = {}

        self.create_widgets()
        self.check_system_files_status()
        self.update_buttons_state()
        self.show()

    def create_tooltip(self, widget, text):
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tk.Label(tooltip, text=text, background="lightyellow",
                     relief="solid", borderwidth=1, font=("Arial", 8)).pack()
            widget.tooltip = tooltip

        def on_leave(event):
            if hasattr(widget, 'tooltip') and widget.tooltip:
                widget.tooltip.destroy()
                widget.tooltip = None

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def create_widgets(self):
        ttk.Label(self.frame, text="Модификация системных DLL-файлов",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)

        ttk.Label(self.frame,
                  text="Данная функция позволяет добавить в систему модифицированные файлы\n"
                       "содержащие иконки от Windows 7. Выберите файлы для модификации:",
                  font=('Arial', 10), foreground='#7f8c8d', justify="center").pack(pady=(0, 20))

        settings_frame = ttk.LabelFrame(self.frame, text="Настройки модификации", padding=15)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.select_folder_button = ttk.Button(
            settings_frame,
            text="Выбрать папку с модифицированными файлами",
            command=self.select_folder, width=40)
        self.select_folder_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.select_folder_button,
                            "Выберите папку, содержащую модифицированные DLL-файлы")

        self.result_label = ttk.Label(settings_frame, text="", foreground="red")
        self.result_label.pack(pady=5)

        checkbox_frame = ttk.Frame(settings_frame)
        checkbox_frame.pack(pady=10, fill=tk.X)

        ttk.Label(checkbox_frame, text="Выберите файлы для модификации:",
                  font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))

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
                file_frame, text=f"{filename} - {description}",
                variable=self.file_vars[filename], command=self.update_buttons_state)
            checkbox.pack(side=tk.LEFT, anchor=tk.W)
            self.create_tooltip(checkbox,
                                f"Включить/выключить модификацию файла {filename}")

            system_status_label = ttk.Label(file_frame, text="", font=('Arial', 8))
            system_status_label.pack(side=tk.RIGHT, padx=5)
            setattr(self, f"{filename.replace('.', '_')}_system_status", system_status_label)

        action_frame = ttk.Frame(settings_frame)
        action_frame.pack(pady=15, fill=tk.X)

        self.start_button = ttk.Button(action_frame, text="Начать",
                                       command=self.start_modification,
                                       state="disabled")
        self.start_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.create_tooltip(self.start_button,
                            "Начать процесс модификации выбранных файлов")

        self.restore_button = ttk.Button(action_frame, text="Восстановить",
                                         command=self.restore_originals,
                                         state="disabled")
        self.restore_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.create_tooltip(self.restore_button,
                            "Восстановить оригинальные версии файлов")

        view_hash_button = ttk.Button(action_frame, text="Посмотреть хеш",
                                      command=self.view_hashes)
        view_hash_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.create_tooltip(view_hash_button,
                            "Показать хеш-суммы файлов в системе")

        ttk.Label(settings_frame,
                  text="⚠️ Внимание: Модификация системных файлов \n"
                       "выполняется на свой страх и риск! \n"
                       "Настоятельно рекомендуется сделать точку восстановления системы",
                  font=('Arial', 8), foreground='#e67e22',
                  justify=tk.CENTER).pack(pady=10)

    def check_system_files_status(self):
        self.system_file_status = {}
        for filename, expected_hash in self.expected_hashes.items():
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            status_label = getattr(self, f"{filename.replace('.', '_')}_system_status")
            try:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        file_hash = hashlib.sha256(f.read()).hexdigest().lower()
                    is_match = file_hash == expected_hash.lower()
                    self.system_file_status[filename] = {
                        'exists': True, 'matches_expected': is_match, 'hash': file_hash}
                    status_label.config(
                        text="✓ В системе" if is_match else "✗ Оригинал",
                        foreground="green" if is_match else "red")
                else:
                    self.system_file_status[filename] = {
                        'exists': False, 'matches_expected': False, 'hash': None}
                    status_label.config(text="✗ Отсутствует", foreground="red")
            except Exception as e:
                self.system_file_status[filename] = {
                    'exists': False, 'matches_expected': False, 'hash': None, 'error': str(e)}
                status_label.config(text="✗ Ошибка", foreground="red")

    def update_buttons_state(self):
        selected_files = [f for f, var in self.file_vars.items() if var.get()]
        if not selected_files:
            self.start_button.config(state="disabled")
            self.restore_button.config(state="disabled")
            return

        any_mismatch = False
        any_match = False
        all_exist = True

        for filename in selected_files:
            status = self.system_file_status.get(filename, {})
            exists = status.get('exists', False)
            matches = status.get('matches_expected', False)
            if not exists:
                all_exist = False
            if matches:
                any_match = True
            else:
                any_mismatch = True

        self.start_button.config(
            state="normal" if (any_mismatch and all_exist and self.selected_folder) else "disabled")
        self.restore_button.config(
            state="normal" if (any_match and all_exist) else "disabled")

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            system32_path = os.path.join(os.environ["WINDIR"], "System32")
            if os.path.normcase(folder_path) == os.path.normcase(system32_path):
                messagebox.showerror("Ошибка", "Выбор папки System32 запрещен!")
                return
            self.selected_folder = folder_path
            self.log_to_file(f"Выбрана папка: {folder_path}")
            self.check_files_and_show_result()
            self.update_buttons_state()

    def check_files_and_show_result(self):
        if not self.selected_folder:
            self.result_label.config(text="Папка не выбрана", foreground="red")
            return

        all_files_exist = True
        all_files_match = True

        for filename, expected_hash in self.expected_hashes.items():
            file_path = os.path.join(self.selected_folder, filename)
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    file_hash = hashlib.sha256(f.read()).hexdigest().lower()
                if file_hash != expected_hash.lower():
                    all_files_match = False
            else:
                all_files_exist = False

        if all_files_exist and all_files_match:
            self.result_label.config(
                text="✓ Все файлы присутствуют и соответствуют эталонным хешам",
                foreground="green")
        elif all_files_exist:
            self.result_label.config(
                text="✓ Все файлы присутствуют, но некоторые не соответствуют хешам",
                foreground="orange")
        else:
            self.result_label.config(
                text="✗ Некоторые файлы отсутствуют или не соответствуют",
                foreground="red")

    def log_to_file(self, message):
        try:
            log_file_path = os.path.join(os.path.expanduser("~"), "Desktop", "mod.log")
            os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(message + "\n")
        except Exception as e:
            print(f"Ошибка записи в лог: {e}")

    def restore_originals(self):
        files_to_process = [f for f, var in self.file_vars.items() if var.get()]
        if not files_to_process:
            messagebox.showwarning("Предупреждение",
                                   "Не выбраны файлы для восстановления")
            return

        files_need_restore = []
        for filename in files_to_process:
            status = self.system_file_status.get(filename, {})
            if status.get('exists') and status.get('matches_expected'):
                files_need_restore.append(filename)

        if not files_need_restore:
            messagebox.showinfo("Информация",
                                "Выбранные файлы уже являются оригинальными "
                                "и не требуют восстановления")
            return

        try:
            self.log_to_file("Завершение процесса explorer.exe...")
            subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при завершении explorer.exe: {e}")

        restored_count = 0
        for filename in files_need_restore:
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            if os.path.exists(file_path):
                try:
                    subprocess.run(f'del "{file_path}" >nul 2>&1', shell=True, check=True)
                except Exception:
                    mod_file_path = f"{file_path}_mod"
                    try:
                        os.rename(file_path, mod_file_path)
                        try:
                            subprocess.run(f'del "{mod_file_path}" >nul 2>&1',
                                           shell=True, check=True)
                        except Exception:
                            pass
                    except Exception as e:
                        self.log_to_file(
                            f"Ошибка при переименовании файла {filename}: {e}")

            orig_file_path = f"{file_path}_orig"
            if os.path.exists(orig_file_path):
                try:
                    os.rename(orig_file_path, file_path)
                    restored_count += 1
                except Exception as e:
                    self.log_to_file(f"Ошибка при восстановлении файла {filename}: {e}")

            acl_file = f"{file_path}.acl"
            if os.path.exists(acl_file):
                try:
                    subprocess.run(f'icacls "{file_path}" /restore "{acl_file}" >nul 2>&1',
                                   shell=True, check=True)
                except Exception:
                    pass

        try:
            subprocess.Popen("explorer.exe", shell=True)
        except Exception:
            pass

        self.check_system_files_status()
        self.update_buttons_state()
        messagebox.showinfo("Успешно",
                            f"Восстановление оригинальных DLL-файлов завершено успешно.\n"
                            f"Восстановлено файлов: {restored_count}")

    def start_modification(self):
        if not self.selected_folder:
            messagebox.showerror("Ошибка",
                                 "Выберите папку с модифицированными файлами")
            return

        files_to_process = [f for f, var in self.file_vars.items() if var.get()]
        if not files_to_process:
            messagebox.showwarning("Предупреждение",
                                   "Не выбраны файлы для модификации")
            return

        files_need_modification = []
        for filename in files_to_process:
            status = self.system_file_status.get(filename, {})
            if status.get('exists') and not status.get('matches_expected'):
                files_need_modification.append(filename)

        if not files_need_modification:
            messagebox.showinfo("Информация",
                                "Выбранные файлы уже соответствуют эталонным хешам "
                                "и не требуют модификации")
            return

        missing_files = [f for f in files_need_modification
                         if not os.path.exists(os.path.join(self.selected_folder, f))]
        if missing_files:
            messagebox.showerror("Ошибка",
                                 f"Файлы отсутствуют в выбранной папке: "
                                 f"{', '.join(missing_files)}!")
            return

        try:
            subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при завершении explorer.exe: {e}")

        modified_count = 0
        for filename in files_need_modification:
            source_file = os.path.join(self.selected_folder, filename)
            target_file = os.path.join(os.environ["WINDIR"], "System32", filename)
            try:
                acl_file = f"{target_file}.acl"
                subprocess.run(f'icacls "{target_file}" /save "{acl_file}" >nul 2>&1',
                               shell=True, check=True)
                subprocess.run(f'takeown /f "{target_file}" /a >nul 2>&1',
                               shell=True, check=True)
                subprocess.run(f'icacls "{target_file}" /grant "%USERNAME%":(F) >nul 2>&1',
                               shell=True, check=True)

                backup_file = f"{target_file}_orig"
                if os.path.exists(target_file):
                    os.rename(target_file, backup_file)

                subprocess.run(f'copy "{source_file}" "{target_file}" >nul 2>&1',
                               shell=True, check=True)
                modified_count += 1

                subprocess.run(f'icacls "{target_file}" /setowner '
                               f'"NT SERVICE\\TrustedInstaller" >nul 2>&1',
                               shell=True, check=True)
                subprocess.run(f'icacls "{target_file}" /restore "{acl_file}" >nul 2>&1',
                               shell=True, check=True)
            except Exception as e:
                self.log_to_file(f"Ошибка при обработке файла {filename}: {e}")

        try:
            subprocess.Popen("explorer.exe", shell=True)
        except Exception as e:
            self.log_to_file(f"Ошибка при запуске explorer.exe: {e}")

        self.check_system_files_status()
        self.update_buttons_state()
        messagebox.showinfo("Успешно",
                            f"Модификация системных DLL файлов завершена успешно.\n"
                            f"Обработано файлов: {modified_count}")

    def view_hashes(self):
        hash_window = tk.Toplevel(self.parent)
        hash_window.title("Хеш-суммы файлов")
        hash_window.geometry("840x200")
        hash_window.resizable(False, False)

        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            try:
                hash_window.iconbitmap(icon_path)
            except Exception:
                pass

        main_frame = ttk.Frame(hash_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(main_frame, text="Хеш-суммы файлов в системе:",
                  font=("Arial", 14, "bold")).pack(anchor=tk.W, pady=(0, 10))

        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(content_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header_frame, text="Файл", font=("Arial", 10, "bold"),
                  width=25).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(header_frame, text="Статус", font=("Arial", 10, "bold"),
                  width=20).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(header_frame, text="SHA-256 Хеш",
                  font=("Arial", 10, "bold")).pack(side=tk.LEFT)

        for filename, expected_hash in self.expected_hashes.items():
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=filename, font=("Courier New", 9),
                      width=25, anchor="w").pack(side=tk.LEFT, padx=(0, 10))

            try:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        file_hash = hashlib.sha256(f.read()).hexdigest().upper()
                    is_match = file_hash == expected_hash.upper()
                    ttk.Label(row_frame,
                              text="✓ Модифицирован" if is_match else "✗ Оригинал",
                              foreground="green" if is_match else "red",
                              font=("Arial", 9), width=20, anchor="w").pack(
                        side=tk.LEFT, padx=(0, 10))
                    ttk.Label(row_frame, text=file_hash, font=("Courier New", 9),
                              anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
                else:
                    ttk.Label(row_frame, text="✗ Не найден", foreground="red",
                              font=("Arial", 9), width=20, anchor="w").pack(
                        side=tk.LEFT, padx=(0, 10))
                    ttk.Label(row_frame, text="Файл не существует",
                              font=("Courier New", 9), foreground="red",
                              anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
            except Exception as e:
                ttk.Label(row_frame, text="✗ Ошибка", foreground="red",
                          font=("Arial", 9), width=20, anchor="w").pack(
                    side=tk.LEFT, padx=(0, 10))
                ttk.Label(row_frame, text=f"Ошибка чтения: {str(e)}",
                          font=("Courier New", 9), foreground="red",
                          anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()