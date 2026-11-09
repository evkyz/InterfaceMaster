import hashlib
import os
import subprocess
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import ttk, messagebox

EXPECTED_HASHES = {
    "imageres.dll": "DDD5D7462A26F7837CD02A7667615354F9377A0A727BF880D825676D6DDC4CB7",
    "imagesp1.dll": "466E64D491C7B3D7051FBDC0907CD813D7A6435D18BF4B00CE0C3191B80E0541",
    "networkexplorer.dll": "7BF68211CFBDDE2E377C992CEC2F5A0C1F41B1F12BE3A99733A549AB63A8C961",
    "DDORes.dll": "4802C8F0BAAE3765FDFF16F09F56647B1C8B20889E4C531B773442EC068CBC3B",
    "mmres.dll": "4651BB26196D0ED493A26F4E401C8FAC0465E7F9DE2ECFED1642155A6EAF3703",
}

FILES_INFO = {
    "imageres.dll": "Основные иконки системы",
    "imagesp1.dll": "Дополнительные иконки",
    "networkexplorer.dll": "Иконки сети и проводника",
    "DDORes.dll": "Иконки DDE",
    "mmres.dll": "Иконки звука",
}

SYSTEM32 = os.path.join(os.environ["WINDIR"], "System32")


def _sha256_file(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest().lower()


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.selected_folder = None
        self.file_vars = {name: tk.BooleanVar(value=False) for name in EXPECTED_HASHES}
        self.start_button = None
        self.restore_button = None
        self.system_file_status = {}
        self.status_labels = {}

        self._create_widgets()
        self._check_system_files_status()
        self._update_buttons_state()
        self.show()

    def _create_tooltip(self, widget, text):
        def enter(event):
            tip = tk.Toplevel()
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tk.Label(tip, text=text, background="lightyellow",
                     relief="solid", borderwidth=1, font=("Arial", 8)).pack()
            widget.tooltip = tip

        def leave(event):
            if getattr(widget, 'tooltip', None):
                widget.tooltip.destroy()
                widget.tooltip = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def _create_widgets(self):
        ttk.Label(self.frame, text="Модификация системных DLL-файлов",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)
        ttk.Label(self.frame,
                  text="Данная функция позволяет добавить в систему модифицированные файлы\n"
                       "содержащие иконки от Windows 7. Выберите файлы для модификации:",
                  font=('Arial', 10), foreground='#7f8c8d',
                  justify="center").pack(pady=(0, 20))

        settings = ttk.LabelFrame(self.frame, text="Настройки модификации", padding=15)
        settings.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        select_btn = ttk.Button(settings,
                                text="Выбрать папку с модифицированными файлами",
                                command=self._select_folder, width=40)
        select_btn.pack(pady=8, fill=tk.X)
        self._create_tooltip(select_btn,
                             "Выберите папку, содержащую модифицированные DLL-файлы")

        self.result_label = ttk.Label(settings, text="", foreground="red")
        self.result_label.pack(pady=5)

        checkbox_frame = ttk.Frame(settings)
        checkbox_frame.pack(pady=10, fill=tk.X)
        ttk.Label(checkbox_frame, text="Выберите файлы для модификации:",
                  font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))

        for filename, description in FILES_INFO.items():
            row = ttk.Frame(checkbox_frame)
            row.pack(fill=tk.X, pady=2)
            cb = ttk.Checkbutton(row, text=f"{filename} - {description}",
                                 variable=self.file_vars[filename],
                                 command=self._update_buttons_state)
            cb.pack(side=tk.LEFT, anchor=tk.W)
            self._create_tooltip(cb, f"Включить/выключить модификацию файла {filename}")
            status = ttk.Label(row, text="", font=('Arial', 8))
            status.pack(side=tk.RIGHT, padx=5)
            self.status_labels[filename] = status

        actions = ttk.Frame(settings)
        actions.pack(pady=15, fill=tk.X)

        self.start_button = ttk.Button(actions, text="Начать",
                                       command=self._start_modification, state="disabled")
        self.start_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self._create_tooltip(self.start_button,
                             "Начать процесс модификации выбранных файлов")

        self.restore_button = ttk.Button(actions, text="Восстановить",
                                         command=self._restore_originals,
                                         state="disabled")
        self.restore_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self._create_tooltip(self.restore_button,
                             "Восстановить оригинальные версии файлов")

        hash_btn = ttk.Button(actions, text="Посмотреть хеш", command=self._view_hashes)
        hash_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self._create_tooltip(hash_btn, "Показать хеш-суммы файлов в системе")

        ttk.Label(settings,
                  text="⚠️ Внимание: Модификация системных файлов\n"
                       "выполняется на свой страх и риск!\n"
                       "Настоятельно рекомендуется сделать точку восстановления системы",
                  font=('Arial', 8), foreground='#e67e22',
                  justify=tk.CENTER).pack(pady=10)

    def _check_system_files_status(self):
        self.system_file_status = {}
        for filename, expected in EXPECTED_HASHES.items():
            path = os.path.join(SYSTEM32, filename)
            status = self.status_labels[filename]
            try:
                if os.path.exists(path):
                    h = _sha256_file(path)
                    match = h == expected.lower()
                    self.system_file_status[filename] = {
                        'exists': True, 'matches_expected': match, 'hash': h}
                    status.config(text="✓ В системе" if match else "✗ Оригинал",
                                  foreground="green" if match else "red")
                else:
                    self.system_file_status[filename] = {
                        'exists': False, 'matches_expected': False, 'hash': None}
                    status.config(text="✗ Отсутствует", foreground="red")
            except Exception as e:
                self.system_file_status[filename] = {
                    'exists': False, 'matches_expected': False, 'hash': None,
                    'error': str(e)}
                status.config(text="✗ Ошибка", foreground="red")

    def _update_buttons_state(self):
        selected = [f for f, v in self.file_vars.items() if v.get()]
        if not selected:
            self.start_button.config(state="disabled")
            self.restore_button.config(state="disabled")
            return

        any_mismatch = any_match = False
        all_exist = True
        for f in selected:
            s = self.system_file_status.get(f, {})
            if not s.get('exists'):
                all_exist = False
            if s.get('matches_expected'):
                any_match = True
            else:
                any_mismatch = True

        self.start_button.config(
            state="normal" if (any_mismatch and all_exist and self.selected_folder)
            else "disabled")
        self.restore_button.config(
            state="normal" if (any_match and all_exist) else "disabled")

    def _select_folder(self):
        path = filedialog.askdirectory()
        if not path:
            return
        if os.path.normcase(path) == os.path.normcase(SYSTEM32):
            messagebox.showerror("Ошибка", "Выбор папки System32 запрещен!")
            return
        self.selected_folder = path
        self._log(f"Выбрана папка: {path}")
        self._check_files_and_show_result()
        self._update_buttons_state()

    def _check_files_and_show_result(self):
        if not self.selected_folder:
            self.result_label.config(text="Папка не выбрана", foreground="red")
            return
        all_exist = all_match = True
        for filename, expected in EXPECTED_HASHES.items():
            path = os.path.join(self.selected_folder, filename)
            if os.path.exists(path):
                if _sha256_file(path) != expected.lower():
                    all_match = False
            else:
                all_exist = False
        if all_exist and all_match:
            self.result_label.config(
                text="✓ Все файлы присутствуют и соответствуют эталонным хешам",
                foreground="green")
        elif all_exist:
            self.result_label.config(
                text="✓ Все файлы присутствуют, но некоторые не соответствуют хешам",
                foreground="orange")
        else:
            self.result_label.config(
                text="✗ Некоторые файлы отсутствуют или не соответствуют",
                foreground="red")

    def _log(self, message):
        try:
            log_path = os.path.join(os.path.expanduser("~"), "Desktop", "mod.log")
            os.makedirs(os.path.dirname(log_path), exist_ok=True)
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(message + "\n")
        except Exception:
            pass

    def _restore_originals(self):
        files = [f for f, v in self.file_vars.items() if v.get()]
        if not files:
            messagebox.showwarning("Предупреждение", "Не выбраны файлы для восстановления")
            return

        to_restore = [f for f in files
                      if self.system_file_status.get(f, {}).get('exists')
                      and self.system_file_status[f].get('matches_expected')]
        if not to_restore:
            messagebox.showinfo("Информация",
                                "Выбранные файлы уже являются оригинальными "
                                "и не требуют восстановления")
            return

        try:
            subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
        except Exception as e:
            self._log(f"Ошибка при завершении explorer.exe: {e}")

        restored = 0
        for filename in to_restore:
            target = os.path.join(SYSTEM32, filename)
            if os.path.exists(target):
                try:
                    subprocess.run(f'del "{target}" >nul 2>&1', shell=True, check=True)
                except Exception:
                    try:
                        os.rename(target, f"{target}_mod")
                        subprocess.run(f'del "{target}_mod" >nul 2>&1', shell=True,
                                       check=True)
                    except Exception as e:
                        self._log(f"Ошибка при переименовании {filename}: {e}")
            orig = f"{target}_orig"
            if os.path.exists(orig):
                try:
                    os.rename(orig, target)
                    restored += 1
                except Exception as e:
                    self._log(f"Ошибка при восстановлении {filename}: {e}")
            acl = f"{target}.acl"
            if os.path.exists(acl):
                try:
                    subprocess.run(f'icacls "{target}" /restore "{acl}" >nul 2>&1',
                                   shell=True, check=True)
                except Exception:
                    pass

        subprocess.Popen("explorer.exe", shell=True)
        self._check_system_files_status()
        self._update_buttons_state()
        messagebox.showinfo("Успешно",
                            f"Восстановление оригинальных DLL-файлов завершено.\n"
                            f"Восстановлено файлов: {restored}")

    def _start_modification(self):
        if not self.selected_folder:
            messagebox.showerror("Ошибка", "Выберите папку с модифицированными файлами")
            return
        files = [f for f, v in self.file_vars.items() if v.get()]
        if not files:
            messagebox.showwarning("Предупреждение", "Не выбраны файлы для модификации")
            return

        to_modify = [f for f in files
                     if self.system_file_status.get(f, {}).get('exists')
                     and not self.system_file_status[f].get('matches_expected')]
        if not to_modify:
            messagebox.showinfo("Информация",
                                "Выбранные файлы уже соответствуют эталонным хешам "
                                "и не требуют модификации")
            return

        missing = [f for f in to_modify
                   if not os.path.exists(os.path.join(self.selected_folder, f))]
        if missing:
            messagebox.showerror("Ошибка",
                                 f"Файлы отсутствуют в выбранной папке: "
                                 f"{', '.join(missing)}!")
            return

        try:
            subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
        except Exception as e:
            self._log(f"Ошибка при завершении explorer.exe: {e}")

        modified = 0
        for filename in to_modify:
            src = os.path.join(self.selected_folder, filename)
            target = os.path.join(SYSTEM32, filename)
            try:
                acl = f"{target}.acl"
                subprocess.run(f'icacls "{target}" /save "{acl}" >nul 2>&1',
                               shell=True, check=True)
                subprocess.run(f'takeown /f "{target}" /a >nul 2>&1',
                               shell=True, check=True)
                subprocess.run(f'icacls "{target}" /grant "%USERNAME%":(F) >nul 2>&1',
                               shell=True, check=True)

                backup = f"{target}_orig"
                if os.path.exists(target):
                    os.rename(target, backup)

                subprocess.run(f'copy "{src}" "{target}" >nul 2>&1', shell=True, check=True)
                modified += 1

                subprocess.run(f'icacls "{target}" /setowner '
                               f'"NT SERVICE\\TrustedInstaller" >nul 2>&1',
                               shell=True, check=True)
                subprocess.run(f'icacls "{target}" /restore "{acl}" >nul 2>&1',
                               shell=True, check=True)
            except Exception as e:
                self._log(f"Ошибка при обработке файла {filename}: {e}")

        subprocess.Popen("explorer.exe", shell=True)
        self._check_system_files_status()
        self._update_buttons_state()
        messagebox.showinfo("Успешно",
                            f"Модификация системных DLL файлов завершена.\n"
                            f"Обработано файлов: {modified}")

    def _view_hashes(self):
        win = tk.Toplevel(self.parent)
        win.title("Хеш-суммы файлов")
        win.geometry("840x200")
        win.resizable(False, False)
        try:
            icon = os.path.join(os.path.dirname(__file__), "icon.ico")
            if os.path.exists(icon):
                win.iconbitmap(icon)
        except Exception:
            pass

        main = ttk.Frame(win)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        ttk.Label(main, text="Хеш-суммы файлов в системе:",
                  font=("Arial", 14, "bold")).pack(anchor=tk.W, pady=(0, 10))

        content = ttk.Frame(main)
        content.pack(fill=tk.BOTH, expand=True)
        header = ttk.Frame(content)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="Файл", font=("Arial", 10, "bold"), width=25).pack(
            side=tk.LEFT, padx=(0, 10))
        ttk.Label(header, text="Статус", font=("Arial", 10, "bold"), width=20).pack(
            side=tk.LEFT, padx=(0, 10))
        ttk.Label(header, text="SHA-256 Хеш",
                  font=("Arial", 10, "bold")).pack(side=tk.LEFT)

        for filename, expected in EXPECTED_HASHES.items():
            path = os.path.join(SYSTEM32, filename)
            row = ttk.Frame(content)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=filename, font=("Courier New", 9), width=25,
                      anchor="w").pack(side=tk.LEFT, padx=(0, 10))
            try:
                if os.path.exists(path):
                    h = _sha256_file(path).upper()
                    match = h == expected.upper()
                    ttk.Label(row,
                              text="✓ Модифицирован" if match else "✗ Оригинал",
                              foreground="green" if match else "red",
                              font=("Arial", 9), width=20, anchor="w").pack(
                        side=tk.LEFT, padx=(0, 10))
                    ttk.Label(row, text=h, font=("Courier New", 9),
                              anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
                else:
                    ttk.Label(row, text="✗ Не найден", foreground="red",
                              font=("Arial", 9), width=20, anchor="w").pack(
                        side=tk.LEFT, padx=(0, 10))
                    ttk.Label(row, text="Файл не существует",
                              font=("Courier New", 9), foreground="red",
                              anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)
            except Exception as e:
                ttk.Label(row, text="✗ Ошибка", foreground="red",
                          font=("Arial", 9), width=20, anchor="w").pack(
                    side=tk.LEFT, padx=(0, 10))
                ttk.Label(row, text=f"Ошибка чтения: {e}",
                          font=("Courier New", 9), foreground="red",
                          anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()