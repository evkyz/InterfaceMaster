import os
import subprocess
import sys
import tkinter as tk
import winreg
from tkinter import ttk, messagebox

HKCR = winreg.HKEY_CLASSES_ROOT
HKCU = winreg.HKEY_CURRENT_USER
HKLM = winreg.HKEY_LOCAL_MACHINE

DESKTOP_SHELL = r'DesktopBackground\Shell'
BLOCKED_PATH = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
COMPUTER_CLSID = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell'

ADMIN_ITEMS = "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run"
SYSTEM_ITEMS = ("control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;defrag;"
                "folderoptions;networksh;appwiz;power;rstrui;search")

COMMAND_STORE = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell'

ADMIN_COMMANDS = {
    'cmd': ('Командная строка', 'cmd.exe', 'cmd.exe'),
    'perfmon': ('Монитор ресурсов', 'imageres.dll,144', 'perfmon.exe /res'),
    'relmon': ('Монитор надежности системы', 'perfmon.exe', 'perfmon.exe /rel'),
    'trouble': ('Устранение неполадок', 'imageres.dll,124', 'control /name Microsoft.Troubleshooting'),
    'services': ('Службы', 'filemgmt.dll', 'mmc.exe services.msc'),
    'event': ('Просмотр событий', 'eventvwr.exe', 'eventvwr.exe'),
    'taskschd': ('Планировщик заданий', 'miguiresource.dll,1', 'Control schedtasks'),
    'useracc': ('Учетные записи', 'imageres.dll,74', 'Control userpasswords'),
    'useracc2': ('Учетные записи (классич.)', 'shell32.dll,111', 'Control userpasswords2'),
    'network': ('Сетевые подключения', 'ncpa.cpl', 'control ncpa.cpl'),
    'wf': ('Брандмауэр Windows', 'wscui.cpl,3', 'wf.msc'),
    'run': ('Выполнить', 'shell32.dll,24',
            'explorer shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'),
}

SYSTEM_COMMANDS = {
    'control': ('Панель управления', 'control.exe', 'control.exe'),
    'sysdm': ('Свойства системы', 'sysdm.cpl', 'control sysdm.cpl'),
    'msconfig': ('Конфигурация системы', 'msconfig.exe', 'msconfig.exe'),
    'regedit': ('Редактор реестра', 'regedit.exe', 'regedit.exe'),
    'taskmgr': ('Диспетчер задач', 'taskmgr.exe', 'taskmgr.exe'),
    'devmgr': ('Диспетчер устройств', 'devmgr.dll,4', 'mmc.exe devmgmt.msc'),
    'cleanmgr': ('Очистка диска', 'cleanmgr.exe', 'cleanmgr.exe'),
    'defrag': ('Дефрагментация диска', 'dfrgui.exe', 'dfrgui.exe'),
    'folderoptions': ('Свойства папки', 'explorer.exe', 'control folders'),
    'networksh': ('Центр управления сетями', 'networkexplorer.dll,4',
                  'control /name Microsoft.NetworkAndSharingCenter'),
    'appwiz': ('Программы и компоненты', 'appwiz.cpl', 'control appwiz.cpl'),
    'power': ('Электропитание', 'powercfg.cpl', 'control powercfg.cpl'),
    'rstrui': ('Восстановление системы', 'rstrui.exe', 'rstrui.exe'),
    'search': ('Поиск', 'shell32.dll,22',
               'explorer shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}'),
}


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.context_menu_module = ContextMenuModule(self.frame)
        self.show()

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()


class ContextMenuModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_new_module = None
        self.sendto_module = None
        try:
            from createnew import CreateNewModule
            from sendto import SendToModule
            self.create_new_module = CreateNewModule
            self.sendto_module = SendToModule
        except ImportError:
            pass
        self.create_widgets()
        self.show()

    def create_widgets(self):
        ttk.Label(self.parent, text="Настройка контекстного меню",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)

        notebook = ttk.Notebook(self.parent)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        desktop_frame = ttk.Frame(notebook)
        create_frame = ttk.Frame(notebook)
        sendto_frame = ttk.Frame(notebook)

        notebook.add(desktop_frame, text="     Меню рабочего стола     ")
        notebook.add(create_frame, text="          Меню Создать          ")
        notebook.add(sendto_frame, text="         Меню Отправить         ")

        self.desktop_menu_tab = DesktopMenuTab(desktop_frame)

        if self.create_new_module:
            self.create_new_tab = self.create_new_module(create_frame)
        else:
            ttk.Label(create_frame, text="Модуль 'Создать' не найден").pack(pady=20)

        if self.sendto_module:
            self.sendto_tab = self.sendto_module(sendto_frame)
        else:
            ttk.Label(sendto_frame, text="Модуль 'Отправить' не найден").pack(pady=20)

    def show(self):
        self.parent.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.parent.pack_forget()


# ---------- registry helpers ----------

def _reg_exists(hive, path):
    try:
        with winreg.OpenKey(hive, path):
            return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


def _delete_key_recursive(hive, path):
    try:
        with winreg.OpenKey(hive, path) as key:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(key, i)
                except OSError:
                    break
                _delete_key_recursive(hive, path + "\\" + sub)
                i += 1
    except FileNotFoundError:
        return
    try:
        winreg.DeleteKey(hive, path)
    except OSError:
        pass


def _copy_key(hive, src, dst):
    winreg.CreateKey(hive, dst)
    with winreg.OpenKey(hive, src) as sk, winreg.OpenKey(hive, dst, 0, winreg.KEY_WRITE) as dk:
        i = 0
        while True:
            try:
                name, value, vtype = winreg.EnumValue(sk, i)
            except OSError:
                break
            winreg.SetValueEx(dk, name, 0, vtype, value)
            i += 1
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(sk, i)
            except OSError:
                break
            _copy_key(hive, src + "\\" + sub, dst + "\\" + sub)
            i += 1


def _rename_key(hive, old, new):
    try:
        _copy_key(hive, old, new)
        _delete_key_recursive(hive, old)
        return True
    except Exception:
        return False


def _run_ps(script):
    """Runs a PowerShell script hidden. Returns returncode."""
    try:
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        r = subprocess.run(["powershell", "-Command", script],
                           capture_output=True, text=True, encoding='cp866',
                           timeout=30, startupinfo=si, shell=True)
        return r.returncode == 0
    except Exception:
        return False


def _ps_escape(s):
    return s.replace("'", "''")


# ---------- Desktop menu tab ----------

class DesktopMenuTab:
    def __init__(self, parent):
        self.parent = parent
        self.menu_item_buttons = []
        self._build()

    # ---- UI ----
    def _build(self):
        canvas = tk.Canvas(self.parent)
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=canvas.yview)
        frame = ttk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        ttk.Label(frame, text="Управление пунктами контекстного меню",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        ctrl = ttk.Frame(frame)
        ctrl.pack(fill=tk.X, padx=10, pady=10)
        btns = ttk.Frame(ctrl)
        btns.pack(expand=True)

        ttk.Button(btns, text="Перезапустить Проводник",
                   command=self.restart_explorer, width=25).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(btns, text="По умолчанию",
                   command=self.set_default_settings, width=15).pack(side=tk.LEFT, padx=5, pady=5)

        for item in self._build_items():
            row = ttk.Frame(frame)
            row.pack(fill=tk.X, padx=10, pady=5)
            btn = self._create_item_button(row, item)
            self.menu_item_buttons.append({"button": btn, "data": item})

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.update_menu_items_status()

    def _build_items(self):
        return [
            {"name": "Изменить с помощью Paint 3D", "add_command": self.add_paint3d_edit,
             "remove_command": self.remove_paint3d_edit,
             "check_command": lambda: self.check_menu_item_exists("paint3d"),
             "default_enabled": True},
            {"name": "Исправление проблем с совместимостью",
             "add_command": self.add_compatibility, "remove_command": self.remove_compatibility,
             "check_command": lambda: self.check_menu_item_exists("compatibility"),
             "default_enabled": True},
            {"name": "Закрепить на начальном экране",
             "add_command": self.add_pin_to_start, "remove_command": self.remove_pin_to_start,
             "check_command": lambda: self.check_menu_item_exists("pin_to_start"),
             "default_enabled": True},
            {"name": "Закрепить на Панели задач",
             "add_command": self.add_pin_to_taskbar, "remove_command": self.remove_pin_to_taskbar,
             "check_command": lambda: self.check_menu_item_exists("pin_to_taskbar"),
             "default_enabled": True},
            {"name": "Копировать как путь",
             "add_command": self.add_copy_as_path, "remove_command": self.remove_copy_as_path,
             "check_command": lambda: self.check_menu_item_exists("copy_as_path"),
             "default_enabled": True},
            {"name": "Отправить (metro)",
             "add_command": self.add_modern_sharing, "remove_command": self.remove_modern_sharing,
             "check_command": lambda: self.check_menu_item_exists("modern_sharing"),
             "default_enabled": True},
            {"name": "Передать на устройство",
             "add_command": self.add_send_to_device, "remove_command": self.remove_send_to_device,
             "check_command": lambda: self.check_menu_item_exists("send_to_device"),
             "default_enabled": True},
            {"name": "Предоставить доступ к",
             "add_command": self.add_share_access, "remove_command": self.remove_share_access,
             "check_command": lambda: self.check_menu_item_exists("share_access"),
             "default_enabled": True},
            {"name": "Проверка с использованием Windows Defender",
             "add_command": self.add_defender_scan, "remove_command": self.remove_defender_scan,
             "check_command": lambda: self.check_menu_item_exists("defender"),
             "default_enabled": True},
            {"name": "Восстановить прежнюю версию",
             "add_command": self.add_previous_version, "remove_command": self.remove_previous_version,
             "check_command": lambda: self.check_menu_item_exists("previous_version"),
             "default_enabled": True},
            {"name": "Стать владельцем",
             "add_command": self.add_take_ownership, "remove_command": self.remove_take_ownership,
             "check_command": lambda: self.check_menu_item_exists("take_ownership"),
             "default_enabled": False},
            {"name": "Копировать/Переместить в папку",
             "add_command": self.add_copy_move_menu, "remove_command": self.remove_copy_move_menu,
             "check_command": lambda: self.check_menu_item_exists("copy_move"),
             "default_enabled": False},
            {"name": "Удалить содержимое папки",
             "add_command": self.add_delete_folder_content,
             "remove_command": self.remove_delete_folder_content,
             "check_command": lambda: self.check_menu_item_exists("delete_folder_content"),
             "default_enabled": False},
            {"name": "Диспетчер задач",
             "add_command": self.add_task_manager, "remove_command": self.remove_task_manager,
             "check_command": lambda: self.check_menu_item_exists("task_manager"),
             "default_enabled": False},
            {"name": "Очистить буфер",
             "add_command": self.add_clipboard_clear, "remove_command": self.remove_clipboard_clear,
             "check_command": lambda: self.check_menu_item_exists("clipboard_clear"),
             "default_enabled": False},
            {"name": "Перезапустить explorer",
             "add_command": self.add_restart_explorer, "remove_command": self.remove_restart_explorer,
             "check_command": lambda: self.check_menu_item_exists("restart_explorer"),
             "default_enabled": False},
            {"name": "Редактор реестра",
             "add_command": self.add_regedit, "remove_command": self.remove_regedit,
             "check_command": lambda: self.check_menu_item_exists("regedit"),
             "default_enabled": False},
            {"name": "Выключение/Перезагрузка/Гибернация",
             "add_command": self.add_power_menu, "remove_command": self.remove_power_menu,
             "check_command": lambda: self.check_menu_item_exists("power_menu"),
             "default_enabled": False},
            {"name": "Администрирование",
             "add_command": self.add_admin_menu, "remove_command": self.remove_admin_menu,
             "check_command": lambda: self.check_menu_item_exists("admin_menu"),
             "default_enabled": False},
            {"name": "Система",
             "add_command": self.add_system_menu, "remove_command": self.remove_system_menu,
             "check_command": lambda: self.check_menu_item_exists("system_menu"),
             "default_enabled": False},
        ]

    def _create_item_button(self, parent, item):
        row = ttk.Frame(parent)
        row.pack()
        canvas = tk.Canvas(row, width=20, height=20, highlightthickness=0)
        canvas.pack(side=tk.LEFT, padx=(0, 5))
        canvas.create_oval(2, 2, 18, 18, outline='#e74c3c', fill='#e74c3c')
        canvas.create_text(10, 10, text="✗", fill="white", font=('Arial', 10, 'bold'))
        item["indicator_canvas"] = canvas
        item["indicator_circle"], item["indicator_text"] = canvas.find_all()

        btn = ttk.Button(row, text=item["name"], width=50,
                         command=lambda i=item: self.toggle_menu_item(i))
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=2)
        return btn

    def _update_indicator(self, item, enabled):
        canvas = item.get('indicator_canvas')
        if not canvas:
            return
        color = '#27ae60' if enabled else '#e74c3c'
        canvas.itemconfig(item['indicator_circle'], outline=color, fill=color)
        canvas.itemconfig(item['indicator_text'], text="✓" if enabled else "✗")

    # ---- actions ----
    def restart_explorer(self):
        try:
            subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True,
                           capture_output=True)
            subprocess.Popen("explorer.exe")
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при перезагрузке проводника")

    def set_default_settings(self, silent=False):
        for entry in self.menu_item_buttons:
            data = entry["data"]
            try:
                current = data["check_command"]()
                want = data.get("default_enabled", True)
                if want and not current:
                    data["add_command"]()
                elif not want and current:
                    data["remove_command"]()
            except Exception:
                continue
        self.update_power_menu_icons()
        self.update_menu_items_status()
        self.update_desktop_shell_order()
        if not silent:
            messagebox.showinfo("Успех", "Настройки сброшены по умолчанию!")

    def toggle_menu_item(self, item):
        try:
            exists = item["check_command"]()
            (item["remove_command"] if exists else item["add_command"])()
            self._update_indicator(item, not exists)
            self.update_menu_items_status()
            if item["name"] in ("Выключение/Перезагрузка/Гибернация", "Диспетчер задач",
                                "Перезапустить explorer", "Очистить буфер", "Редактор реестра"):
                self.update_desktop_shell_order()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта: {e}")

    def update_menu_items_status(self):
        for entry in self.menu_item_buttons:
            item = entry["data"]
            try:
                self._update_indicator(item, item["check_command"]())
            except Exception:
                pass

    # ---- shared desktop-shell helpers ----
    def update_desktop_shell_order(self):
        try:
            active = []
            for key, name in (("task_manager", "TaskManager"),
                              ("restart_explorer", "RestartExplorer"),
                              ("clipboard_clear", "Clip"),
                              ("regedit", "Regedit")):
                if self.check_menu_item_exists(key):
                    active.append(name)
            if self.check_menu_item_exists("power_menu"):
                for p, n in ((r'DesktopBackground\Shell\Shutdown', "Shutdown"),
                             (r'DesktopBackground\Shell\Reboot', "Reboot"),
                             (r'DesktopBackground\Shell\Hibernation', "Hibernation")):
                    if _reg_exists(HKCR, p):
                        active.append(n)
            active += ["Display", "Personalize"]
            winreg.CreateKey(HKCR, DESKTOP_SHELL)
            with winreg.OpenKey(HKCR, DESKTOP_SHELL, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, ",".join(active))
        except Exception:
            pass

    def _create_desktop_shell_item(self, key_path, mui_verb, icon, command,
                                    extended=False):
        """Создаёт пункт контекстного меню рабочего стола.

        extended=False (по умолчанию) — пункт виден всегда.
        extended=True — пункт виден только с зажатым Shift.
        """
        try:
            winreg.CreateKey(HKCR, key_path)
            with winreg.OpenKey(HKCR, key_path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, mui_verb)
                if extended:
                    winreg.SetValueEx(k, "Extended", 0, winreg.REG_SZ, "")
                else:
                    try:
                        winreg.DeleteValue(k, "Extended")
                    except FileNotFoundError:
                        pass
                if icon:
                    winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon)
            cmd_key = key_path + r'\command'
            winreg.CreateKey(HKCR, cmd_key)
            with winreg.OpenKey(HKCR, cmd_key, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, command)
            self.update_desktop_shell_order()
            self.update_menu_items_status()
        except PermissionError:
            messagebox.showerror("Ошибка", f"Нет прав для '{mui_verb}'")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")

    def _delete_desktop_shell_item(self, key_path):
        try:
            try:
                winreg.DeleteKey(HKCR, key_path + r'\command')
            except FileNotFoundError:
                pass
            winreg.DeleteKey(HKCR, key_path)
        except FileNotFoundError:
            pass
        except PermissionError:
            messagebox.showerror("Ошибка", "Нет прав для удаления пункта")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")
        self.update_desktop_shell_order()
        self.update_menu_items_status()

    # ---- icon path helper ----
    @staticmethod
    def check_imaster_dll_from_main():
        try:
            import __main__ as main
            if not hasattr(main, "DLLChecker"):
                return None
            path = main.DLLChecker.get_dll_path_without_prompt()
            if not path or not os.path.isfile(path):
                return None
            meipass = getattr(sys, "_MEIPASS", None)
            if meipass:
                try:
                    if os.path.commonpath([os.path.abspath(path), os.path.abspath(meipass)]) == os.path.abspath(meipass):
                        return None
                except ValueError:
                    pass
            return path
        except Exception:
            return None

    def update_power_menu_icons(self):
        path = self.check_imaster_dll_from_main()
        mapping = {r'DesktopBackground\Shell\Reboot': -1001,
                   r'DesktopBackground\Shell\Hibernation': -1002,
                   r'DesktopBackground\Shell\Clip': -1003}
        for key_path, icon_idx in mapping.items():
            try:
                with winreg.OpenKey(HKCR, key_path, 0, winreg.KEY_SET_VALUE) as k:
                    if path:
                        winreg.SetValueEx(k, "Icon", 0, winreg.REG_EXPAND_SZ, f"{path},{icon_idx}")
                    else:
                        try:
                            winreg.DeleteValue(k, "Icon")
                        except FileNotFoundError:
                            pass
            except FileNotFoundError:
                pass
            except Exception:
                pass

    # ---- admin / system menus ----
    def _create_command_store_entries(self, commands):
        try:
            for name, (display, icon, command) in commands.items():
                try:
                    path = f"{COMMAND_STORE}\\{name}"
                    winreg.CreateKey(HKLM, path)
                    with winreg.OpenKey(HKLM, path, 0, winreg.KEY_SET_VALUE) as k:
                        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, display)
                        winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon)
                    cmd_path = path + "\\command"
                    winreg.CreateKey(HKLM, cmd_path)
                    with winreg.OpenKey(HKLM, cmd_path, 0, winreg.KEY_SET_VALUE) as k:
                        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, command)
                except Exception:
                    continue
            return True
        except Exception:
            return False

    def _create_clsid_menu(self, suffix, verb, items, icon):
        try:
            self._create_command_store_entries(ADMIN_COMMANDS if suffix == "Admin" else SYSTEM_COMMANDS)
            path = f"{COMPUTER_CLSID}\\{suffix}"
            winreg.CreateKey(HKCU, path)
            with winreg.OpenKey(HKCU, path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, verb)
                winreg.SetValueEx(k, "SubCommands", 0, winreg.REG_SZ, items)
                winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, icon)
                winreg.SetValueEx(k, "Position", 0, winreg.REG_SZ, "Top")
            self.update_menu_items_status()
            return True
        except Exception:
            return False

    def _remove_clsid_menu(self, suffix):
        for hive, base in ((HKCU, 'Software\\Classes\\CLSID'),
                           (HKLM, 'SOFTWARE\\Classes\\CLSID')):
            _delete_key_recursive(
                hive, f'{base}\\{{20D04FE0-3AEA-1069-A2D8-08002B30309D}}\\shell\\{suffix}')
        self.update_menu_items_status()

    def add_admin_menu(self):
        if not self._create_clsid_menu("Admin", "Администрирование", ADMIN_ITEMS, "mmc.exe"):
            messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Администрирование'")

    def remove_admin_menu(self):
        self._remove_clsid_menu("Admin")

    def add_system_menu(self):
        if not self._create_clsid_menu("System", "Система", SYSTEM_ITEMS, "imageres.dll,104"):
            messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Система'")

    def remove_system_menu(self):
        self._remove_clsid_menu("System")

    # ---- desktop menu simple items ----
    def add_regedit(self):
        # Редактор реестра — скрыт за Shift
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\Regedit',
            "Редактор реестра", "regedit.exe,0", "regedit",
            extended=True)

    def remove_regedit(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\Regedit')

    # ---- VBS helpers ----
    def _restart_explorer_vbs_path(self):
        return os.path.join(os.environ['WINDIR'], 'System32',
                            'hidden_restart_explorer.vbs')

    def _ensure_restart_explorer_vbs(self):
        """Создаёт VBScript для скрытого перезапуска explorer без окна cmd."""
        vbs_path = self._restart_explorer_vbs_path()
        content = (
            'Set WshShell = CreateObject("WScript.Shell")\n'
            'WshShell.Run "cmd.exe /c taskkill /f /im explorer.exe & start explorer.exe", 0, False\n'
        )
        try:
            with open(vbs_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Не удалось создать вспомогательный скрипт:\n{e}")
            return False

    def _clear_clipboard_vbs_path(self):
        return os.path.join(os.environ['WINDIR'], 'System32',
                            'hidden_clear_clipboard.vbs')

    def _ensure_clear_clipboard_vbs(self):
        """Создаёт VBScript для очистки буфера без окна cmd."""
        vbs_path = self._clear_clipboard_vbs_path()
        content = (
            'Set WshShell = CreateObject("WScript.Shell")\n'
            'WshShell.Run "cmd.exe /c echo off | clip", 0, False\n'
        )
        try:
            with open(vbs_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка",
                                 f"Не удалось создать вспомогательный скрипт:\n{e}")
            return False

    def add_restart_explorer(self):
        # Перезапустить explorer — скрыт за Shift, без окна cmd через VBS
        if not self._ensure_restart_explorer_vbs():
            return
        vbs_path = self._restart_explorer_vbs_path()
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\RestartExplorer',
            "Перезапустить explorer",
            "explorer.exe",
            f'wscript.exe //B //Nologo "{vbs_path}"',
            extended=True)

    def remove_restart_explorer(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\RestartExplorer')
        vbs_path = self._restart_explorer_vbs_path()
        try:
            if os.path.exists(vbs_path):
                os.remove(vbs_path)
        except Exception:
            pass

    def add_task_manager(self):
        # Диспетчер задач — виден всегда
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\TaskManager',
            "Диспетчер задач", "taskmgr.exe,0", "taskmgr.exe",
            extended=False)

    def remove_task_manager(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\TaskManager')

    def add_clipboard_clear(self):
        # Очистить буфер — скрыт за Shift, без окна cmd через VBS
        if not self._ensure_clear_clipboard_vbs():
            return
        dll = self.check_imaster_dll_from_main()
        vbs_path = self._clear_clipboard_vbs_path()
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\Clip', "Очистить буфер",
            f"{dll},-1003" if dll else None,
            f'wscript.exe //B //Nologo "{vbs_path}"',
            extended=True)
        self.update_power_menu_icons()

    def remove_clipboard_clear(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\Clip')
        vbs_path = self._clear_clipboard_vbs_path()
        try:
            if os.path.exists(vbs_path):
                os.remove(vbs_path)
        except Exception:
            pass

    def add_power_menu(self):
        # Выключение и Перезагрузка — видимы всегда.
        # Гибернация — скрыта за Shift.
        dll = self.check_imaster_dll_from_main()
        try:
            for key, verb, icon, cmd, ext in (
                (r'DesktopBackground\Shell\Hibernation', "Гибернация",
                 f"{dll},-1002" if dll else None, "shutdown /h", True),
                (r'DesktopBackground\Shell\Reboot', "Перезагрузка компьютера",
                 f"{dll},-1001" if dll else None, "shutdown /r /t 0", False),
                (r'DesktopBackground\Shell\Shutdown', "Выключение компьютера",
                 "%SystemRoot%\\System32\\shell32.dll,-28", "shutdown /s /t 0", False),
            ):
                winreg.CreateKey(HKCR, key)
                with winreg.OpenKey(HKCR, key, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "Position", 0, winreg.REG_SZ, "Bottom")
                    winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, verb)
                    if ext:
                        winreg.SetValueEx(k, "Extended", 0, winreg.REG_SZ, "")
                    else:
                        try:
                            winreg.DeleteValue(k, "Extended")
                        except FileNotFoundError:
                            pass
                    if icon:
                        winreg.SetValueEx(k, "Icon", 0, winreg.REG_EXPAND_SZ, icon)
                cmd_key = key + r'\command'
                winreg.CreateKey(HKCR, cmd_key)
                with winreg.OpenKey(HKCR, cmd_key, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, cmd)
            self.update_desktop_shell_order()
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении меню питания: {e}")

    def remove_power_menu(self):
        for key in (r'DesktopBackground\Shell\Hibernation\command',
                    r'DesktopBackground\Shell\Hibernation',
                    r'DesktopBackground\Shell\Reboot\Command',
                    r'DesktopBackground\Shell\Reboot',
                    r'DesktopBackground\Shell\Shutdown\Command',
                    r'DesktopBackground\Shell\Shutdown'):
            try:
                winreg.DeleteKey(HKCR, key)
            except OSError:
                pass
        self.update_desktop_shell_order()
        self.update_menu_items_status()

    # ---- simple HKLM value/delete helpers ----
    def _simple_lm_value(self, key_path, value):
        try:
            winreg.CreateKey(HKLM, key_path)
            with winreg.OpenKey(HKLM, key_path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, value)
        except PermissionError:
            messagebox.showerror("Ошибка", "Не удалось получить права")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {e}")
        self.update_menu_items_status()

    def _simple_lm_delete(self, key_path):
        try:
            winreg.DeleteKey(HKLM, key_path)
        except FileNotFoundError:
            pass
        except PermissionError:
            messagebox.showerror("Ошибка", "Не удалось получить права")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {e}")
        self.update_menu_items_status()

    # ---- blocked-value helpers ----
    def _unblock(self, guid):
        try:
            winreg.CreateKey(HKLM, BLOCKED_PATH)
            with winreg.OpenKey(HKLM, BLOCKED_PATH, 0, winreg.KEY_SET_VALUE) as k:
                try:
                    winreg.DeleteValue(k, guid)
                except FileNotFoundError:
                    pass
            self.update_menu_items_status()
        except PermissionError:
            messagebox.showerror("Ошибка", "Не удалось получить права")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {e}")

    def _block(self, guid, description):
        try:
            winreg.CreateKey(HKLM, BLOCKED_PATH)
            with winreg.OpenKey(HKLM, BLOCKED_PATH, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, guid, 0, winreg.REG_SZ, description)
            self.update_menu_items_status()
        except PermissionError:
            messagebox.showerror("Ошибка", "Не удалось получить права")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {e}")

    # ---- copy-as-path / pins ----
    def add_copy_as_path(self):
        self._simple_lm_value(
            r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu',
            "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}")

    def remove_copy_as_path(self):
        self._simple_lm_delete(
            r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu')

    def add_pin_to_start(self):
        self._simple_lm_value(
            r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen',
            "{470C0EBD-5D73-4d58-9CED-E91E22E23282}")
        self._simple_lm_value(
            r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\PintoStartScreen',
            "{470C0EBD-5D73-4d58-9CED-E91E22E23282}")

    def remove_pin_to_start(self):
        self._simple_lm_delete(
            r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen')
        self._simple_lm_delete(
            r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\PintoStartScreen')

    def add_pin_to_taskbar(self):
        self._simple_lm_value(
            r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}',
            "Taskband Pin")

    def remove_pin_to_taskbar(self):
        self._simple_lm_delete(
            r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}')

    # ---- blocked items ----
    def add_send_to_device(self):
        self._unblock("{7AD84985-87B4-4a16-BE58-8B72A5B390F7}")

    def remove_send_to_device(self):
        self._block("{7AD84985-87B4-4a16-BE58-8B72A5B390F7}", "Передать на устройство")

    def add_compatibility(self):
        self._unblock("{1d27f844-3a1f-4410-85ac-14651078412d}")

    def remove_compatibility(self):
        self._block("{1d27f844-3a1f-4410-85ac-14651078412d}",
                    "Исправление проблем с совместимостью")

    def add_previous_version(self):
        self._unblock("{596AB062-B4D2-4215-9F74-E9109B0A8153}")

    def remove_previous_version(self):
        self._block("{596AB062-B4D2-4215-9F74-E9109B0A8153}",
                    "Восстановить прежнюю версию")

    def add_share_access(self):
        self._unblock("{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}")

    def remove_share_access(self):
        self._block("{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}", "")

    # ---- defender: через Shell Extensions\Blocked ----
    _DEFENDER_GUID = "{09A47860-11B0-4DA5-AFA5-26D86198A780}"

    def add_defender_scan(self):
        self._unblock(self._DEFENDER_GUID)

    def remove_defender_scan(self):
        self._block(self._DEFENDER_GUID,
                    "Проверка с использованием Windows Defender")

    # ---- paint3d ----
    def add_paint3d_edit(self):
        for ext in ('bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                    'obj', 'ply', 'png', 'stl', 'tif', 'tiff'):
            try:
                base = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{ext}\\Shell"
                winreg.CreateKey(HKLM, base)
                p3d = f"{base}\\3D Edit"
                winreg.CreateKey(HKLM, p3d)
                with winreg.OpenKey(HKLM, p3d, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ,
                                      "@%SystemRoot%\\system32\\mspaint.exe,-59500")
                cmd_key = f"{p3d}\\command"
                winreg.CreateKey(HKLM, cmd_key)
                with winreg.OpenKey(HKLM, cmd_key, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_EXPAND_SZ,
                                      '%SystemRoot%\\system32\\mspaint.exe "%1" /ForceBootstrapPaint3D')
            except Exception:
                continue
        self.update_menu_items_status()

    def remove_paint3d_edit(self):
        for ext in ('bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                    'obj', 'ply', 'png', 'stl', 'tif', 'tiff'):
            p3d = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{ext}\\Shell\\3D Edit"
            for key in (p3d + "\\command", p3d):
                try:
                    winreg.DeleteKey(HKLM, key)
                except OSError:
                    pass
        self.update_menu_items_status()

    # ---- modern sharing ----
    _MODERN_SHARING_PATHS = (
        r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
        r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
        r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
        r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
        r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing',
    )

    def add_modern_sharing(self):
        clsid = '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'
        for path in self._MODERN_SHARING_PATHS:
            try:
                winreg.CreateKey(HKLM, path)
                with winreg.OpenKey(HKLM, path, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, clsid)
            except Exception:
                continue
        self.update_menu_items_status()

    def remove_modern_sharing(self):
        for path in self._MODERN_SHARING_PATHS:
            try:
                winreg.DeleteKey(HKLM, path)
            except OSError:
                pass
        self.update_menu_items_status()

    # ---- copy/move, delete folder content, take ownership ----
    def add_copy_move_menu(self):
        for path, clsid in (
            (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
             "{C2FBB630-2971-11D1-A18C-00C04FD75D13}"),
            (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To',
             "{C2FBB631-2971-11D1-A18C-00C04FD75D13}"),
        ):
            try:
                winreg.CreateKey(HKLM, path)
                with winreg.OpenKey(HKLM, path, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, clsid)
            except Exception:
                continue
        self.update_menu_items_status()

    def remove_copy_move_menu(self):
        for path in (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
                     r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To'):
            try:
                winreg.DeleteKey(HKLM, path)
            except OSError:
                pass
        self.update_menu_items_status()

    def add_delete_folder_content(self):
        path = r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'
        try:
            winreg.CreateKey(HKLM, path)
            with winreg.OpenKey(HKLM, path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "Удалить содержимое папки")
                winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, "shell32.dll,-240")
            cmd_key = path + r'\command'
            winreg.CreateKey(HKLM, cmd_key)
            with winreg.OpenKey(HKLM, cmd_key, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, "", 0, winreg.REG_SZ,
                                  'cmd.exe /c "cd /d \\"%1\\" && del /q /f /s *.*"')
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {e}")

    def remove_delete_folder_content(self):
        for p in (r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent\command',
                  r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'):
            try:
                winreg.DeleteKey(HKLM, p)
            except OSError:
                pass
        self.update_menu_items_status()

    def add_take_ownership(self):
        try:
            for base_key, cmd in (
                (r'*\shell\runas',
                 'cmd.exe /c takeown /f "%1" && icacls "%1" /grant администраторы:F'),
                (r'Directory\shell\runas',
                 'cmd.exe /c takeown /f "%1" /r /d y && icacls "%1" /grant администраторы:F /t'),
            ):
                winreg.CreateKey(HKCR, base_key)
                with winreg.OpenKey(HKCR, base_key, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "Стать владельцем")
                    winreg.SetValueEx(k, "NoWorkingDirectory", 0, winreg.REG_SZ, "")
                    winreg.SetValueEx(k, "Icon", 0, winreg.REG_SZ, "imageres.dll,54")
                    winreg.SetValueEx(k, "MUIVerb", 0, winreg.REG_SZ, "Стать владельцем")
                cmd_key = base_key + r'\command'
                winreg.CreateKey(HKCR, cmd_key)
                with winreg.OpenKey(HKCR, cmd_key, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.SetValueEx(k, "", 0, winreg.REG_SZ, cmd)
                    winreg.SetValueEx(k, "IsolatedCommand", 0, winreg.REG_SZ, cmd)
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {e}")

    def remove_take_ownership(self):
        for k in (r'*\shell\runas\command', r'*\shell\runas',
                  r'Directory\shell\runas\command', r'Directory\shell\runas'):
            try:
                winreg.DeleteKey(HKCR, k)
            except OSError:
                pass
        self.update_menu_items_status()

    # ---- existence checks ----
    def check_menu_item_exists(self, item_type):
        try:
            if item_type in ("admin_menu", "system_menu"):
                suffix = "Admin" if item_type == "admin_menu" else "System"
                for hive, base in ((HKCU, 'Software\\Classes\\CLSID'),
                                   (HKLM, 'SOFTWARE\\Classes\\CLSID')):
                    path = f'{base}\\{{20D04FE0-3AEA-1069-A2D8-08002B30309D}}\\shell\\{suffix}'
                    if _reg_exists(hive, path):
                        try:
                            with winreg.OpenKey(hive, path) as k:
                                winreg.QueryValueEx(k, "MUIVerb")
                                return True
                        except OSError:
                            continue
                return False

            desktop = {
                "regedit": (r'DesktopBackground\Shell\Regedit', "Редактор реестра"),
                "restart_explorer": (r'DesktopBackground\Shell\RestartExplorer', "Перезапустить explorer"),
                "clipboard_clear": (r'DesktopBackground\Shell\Clip', "Очистить буфер"),
                "task_manager": (r'DesktopBackground\Shell\TaskManager', "Диспетчер задач"),
            }
            if item_type in desktop:
                path, verb = desktop[item_type]
                try:
                    with winreg.OpenKey(HKCR, path) as k:
                        value, _ = winreg.QueryValueEx(k, "MUIVerb")
                        return value == verb
                except FileNotFoundError:
                    return False

            if item_type == "power_menu":
                return any(_reg_exists(HKCR, p) for p in
                           (r'DesktopBackground\Shell\Hibernation',
                            r'DesktopBackground\Shell\Reboot',
                            r'DesktopBackground\Shell\Shutdown'))

            simple = {
                "copy_as_path": (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu',
                                 "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}"),
                "pin_to_taskbar": (
                    r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}',
                    "Taskband Pin"),
            }
            if item_type in simple:
                path, expected = simple[item_type]
                try:
                    with winreg.OpenKey(HKLM, path) as k:
                        value, _ = winreg.QueryValueEx(k, "")
                        return value == expected
                except FileNotFoundError:
                    return False

            if item_type == "pin_to_start":
                expected = "{470C0EBD-5D73-4d58-9CED-E91E22E23282}"
                paths = (
                    r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen',
                    r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\PintoStartScreen',
                )
                for path in paths:
                    try:
                        with winreg.OpenKey(HKLM, path) as k:
                            value, _ = winreg.QueryValueEx(k, "")
                            if value == expected:
                                return True
                    except FileNotFoundError:
                        continue
                return False

            blocked = {
                "send_to_device": "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}",
                "compatibility": "{1d27f844-3a1f-4410-85ac-14651078412d}",
                "previous_version": "{596AB062-B4D2-4215-9F74-E9109B0A8153}",
                "share_access": "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}",
                "defender": self._DEFENDER_GUID,
            }
            if item_type in blocked:
                if not _reg_exists(HKLM, BLOCKED_PATH):
                    return True
                try:
                    with winreg.OpenKey(HKLM, BLOCKED_PATH) as k:
                        try:
                            winreg.QueryValueEx(k, blocked[item_type])
                            return False
                        except FileNotFoundError:
                            return True
                except OSError:
                    return True

            if item_type == "modern_sharing":
                for path in self._MODERN_SHARING_PATHS:
                    if not _reg_exists(HKLM, path):
                        continue
                    try:
                        with winreg.OpenKey(HKLM, path) as k:
                            value, _ = winreg.QueryValueEx(k, "")
                            if value == "{e2bf9676-5f8f-435c-97eb-11607a5bedf7}":
                                return True
                    except OSError:
                        continue
                return False

            if item_type == "paint3d":
                for ext in ('bmp', 'fbx', 'gif'):
                    if _reg_exists(
                            HKLM, f"SOFTWARE\\Classes\\SystemFileAssociations\\.{ext}\\Shell\\3D Edit"):
                        return True
                return False

            if item_type == "delete_folder_content":
                return _reg_exists(
                    HKLM, r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent')

            if item_type == "copy_move":
                return _reg_exists(
                    HKLM, r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To')

            if item_type == "take_ownership":
                return _reg_exists(HKCR, r'*\shell\runas')
        except Exception:
            return False
        return False