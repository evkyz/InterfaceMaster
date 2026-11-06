import os
import subprocess
import tkinter as tk
import winreg
from tkinter import ttk, messagebox
import sys


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

        try:
            from createnew import CreateNewModule
            from sendto import SendToModule
            self.create_new_module = CreateNewModule
            self.sendto_module = SendToModule
        except ImportError:
            self.create_new_module = None
            self.sendto_module = None

        self.create_widgets()
        self.show()

    def create_widgets(self):
        ttk.Label(self.parent, text="Настройка контекстного меню",
                  font=('Arial', 16, 'bold'), foreground='#2c3e50').pack(pady=10)

        notebook = ttk.Notebook(self.parent)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        desktop_menu_frame = ttk.Frame(notebook)
        create_menu_frame = ttk.Frame(notebook)
        sendto_menu_frame = ttk.Frame(notebook)

        notebook.add(desktop_menu_frame, text="     Меню рабочего стола     ")
        notebook.add(create_menu_frame, text="          Меню Создать          ")
        notebook.add(sendto_menu_frame, text="         Меню Отправить         ")

        self.desktop_menu_tab = DesktopMenuTab(desktop_menu_frame)

        if self.create_new_module:
            self.create_new_tab = self.create_new_module(create_menu_frame)
        else:
            ttk.Label(create_menu_frame, text="Модуль 'Создать' не найден").pack(pady=20)

        if self.sendto_module:
            self.sendto_tab = self.sendto_module(sendto_menu_frame)
        else:
            ttk.Label(sendto_menu_frame, text="Модуль 'Отправить' не найден").pack(pady=20)

    def show(self):
        self.parent.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.parent.pack_forget()


class DesktopMenuTab:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        canvas = tk.Canvas(self.parent)
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        ttk.Label(scrollable_frame, text="Управление пунктами контекстного меню",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        control_frame = ttk.Frame(scrollable_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)

        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        ttk.Button(buttons_frame, text="Перезапустить Проводник",
                   command=self.restart_explorer, width=25).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(buttons_frame, text="По умолчанию",
                   command=self.set_default_settings, width=15).pack(side=tk.LEFT, padx=5, pady=5)

        self.menu_items_data = [
            {
                "name": "Изменить с помощью Paint 3D",
                "add_command": self.add_paint3d_edit,
                "remove_command": self.remove_paint3d_edit,
                "check_command": lambda: self.check_menu_item_exists("paint3d"),
                "default_enabled": True
            },
            {
                "name": "Исправление проблем с совместимостью",
                "add_command": self.add_compatibility,
                "remove_command": self.remove_compatibility,
                "check_command": lambda: self.check_menu_item_exists("compatibility"),
                "default_enabled": True
            },
            {
                "name": "Закрепить на начальном экране",
                "add_command": self.add_pin_to_start,
                "remove_command": self.remove_pin_to_start,
                "check_command": lambda: self.check_menu_item_exists("pin_to_start"),
                "default_enabled": True
            },
            {
                "name": "Закрепить на Панели задач",
                "add_command": self.add_pin_to_taskbar,
                "remove_command": self.remove_pin_to_taskbar,
                "check_command": lambda: self.check_menu_item_exists("pin_to_taskbar"),
                "default_enabled": True
            },
            {
                "name": "Копировать как путь",
                "add_command": self.add_copy_as_path,
                "remove_command": self.remove_copy_as_path,
                "check_command": lambda: self.check_menu_item_exists("copy_as_path"),
                "default_enabled": True
            },
            {
                "name": "Отправить (metro)",
                "add_command": self.add_modern_sharing,
                "remove_command": self.remove_modern_sharing,
                "check_command": lambda: self.check_menu_item_exists("modern_sharing"),
                "default_enabled": True
            },
            {
                "name": "Передать на устройство",
                "add_command": self.add_send_to_device,
                "remove_command": self.remove_send_to_device,
                "check_command": lambda: self.check_menu_item_exists("send_to_device"),
                "default_enabled": True
            },
            {
                "name": "Предоставить доступ к",
                "add_command": self.add_share_access,
                "remove_command": self.remove_share_access,
                "check_command": lambda: self.check_menu_item_exists("share_access"),
                "default_enabled": True
            },
            {
                "name": "Проверка с использованием Windows Defender",
                "add_command": self.add_defender_scan,
                "remove_command": self.remove_defender_scan,
                "check_command": lambda: self.check_menu_item_exists("defender"),
                "default_enabled": True
            },
            {
                "name": "Восстановить прежнюю версию",
                "add_command": self.add_previous_version,
                "remove_command": self.remove_previous_version,
                "check_command": lambda: self.check_menu_item_exists("previous_version"),
                "default_enabled": True
            },
            {
                "name": "Стать владельцем",
                "add_command": self.add_take_ownership,
                "remove_command": self.remove_take_ownership,
                "check_command": lambda: self.check_menu_item_exists("take_ownership"),
                "default_enabled": False
            },
            {
                "name": "Копировать/Переместить в папку",
                "add_command": self.add_copy_move_menu,
                "remove_command": self.remove_copy_move_menu,
                "check_command": lambda: self.check_menu_item_exists("copy_move"),
                "default_enabled": False
            },
            {
                "name": "Удалить содержимое папки",
                "add_command": self.add_delete_folder_content,
                "remove_command": self.remove_delete_folder_content,
                "check_command": lambda: self.check_menu_item_exists("delete_folder_content"),
                "default_enabled": False
            },
            {
                "name": "Диспетчер задач",
                "add_command": self.add_task_manager,
                "remove_command": self.remove_task_manager,
                "check_command": lambda: self.check_menu_item_exists("task_manager"),
                "default_enabled": False
            },
            {
                "name": "Очистить буфер",
                "add_command": self.add_clipboard_clear,
                "remove_command": self.remove_clipboard_clear,
                "check_command": lambda: self.check_menu_item_exists("clipboard_clear"),
                "default_enabled": False
            },
            {
                "name": "Перезапустить explorer",
                "add_command": self.add_restart_explorer,
                "remove_command": self.remove_restart_explorer,
                "check_command": lambda: self.check_menu_item_exists("restart_explorer"),
                "default_enabled": False
            },
            {
                "name": "Редактор реестра",
                "add_command": self.add_regedit,
                "remove_command": self.remove_regedit,
                "check_command": lambda: self.check_menu_item_exists("regedit"),
                "default_enabled": False
            },
            {
                "name": "Выключение/Перезагрузка/Гибернация",
                "add_command": self.add_power_menu,
                "remove_command": self.remove_power_menu,
                "check_command": lambda: self.check_menu_item_exists("power_menu"),
                "default_enabled": False
            },
            {
                "name": "Администрирование",
                "add_command": self.add_admin_menu,
                "remove_command": self.remove_admin_menu,
                "check_command": lambda: self.check_menu_item_exists("admin_menu"),
                "default_enabled": False
            },
            {
                "name": "Система",
                "add_command": self.add_system_menu,
                "remove_command": self.remove_system_menu,
                "check_command": lambda: self.check_menu_item_exists("system_menu"),
                "default_enabled": False
            }
        ]

        self.menu_item_buttons = []
        for item in self.menu_items_data:
            button_frame = ttk.Frame(scrollable_frame)
            button_frame.pack(fill=tk.X, padx=10, pady=5)
            btn = self.create_menu_item_button(button_frame, item)
            self.menu_item_buttons.append({"button": btn, "data": item})

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.update_menu_items_status()

    def update_desktop_shell_order(self):
        try:
            shell_key = r'DesktopBackground\Shell'
            active_items = []

            if self.check_menu_item_exists("task_manager"):
                active_items.append("TaskManager")
            if self.check_menu_item_exists("restart_explorer"):
                active_items.append("RestartExplorer")
            if self.check_menu_item_exists("clipboard_clear"):
                active_items.append("Clip")
            if self.check_menu_item_exists("regedit"):
                active_items.append("Regedit")
            if self.check_menu_item_exists("power_menu"):
                for key_path, name in [
                    (r'DesktopBackground\Shell\Shutdown', "Shutdown"),
                    (r'DesktopBackground\Shell\Reboot', "Reboot"),
                    (r'DesktopBackground\Shell\Hibernation', "Hibernation")
                ]:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                            active_items.append(name)
                    except FileNotFoundError:
                        pass

            active_items.append("Display")
            active_items.append("Personalize")

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shell_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, ",".join(active_items))
        except Exception:
            try:
                winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shell_key)
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shell_key, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Display,Personalize")
            except Exception:
                pass

    def create_command_store_entries(self):
        try:
            command_store_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell'
            entries = {
                'cmd': ('Командная строка', 'cmd.exe', 'cmd.exe'),
                'perfmon': ('Монитор ресурсов', 'imageres.dll,144', 'perfmon.exe /res'),
                'relmon': ('Монитор надежности системы', 'perfmon.exe', 'perfmon.exe /rel'),
                'trouble': ('Устранение неполадок', 'imageres.dll,124',
                            'control /name Microsoft.Troubleshooting'),
                'services': ('Службы', 'filemgmt.dll', 'mmc.exe services.msc'),
                'event': ('Просмотр событий', 'eventvwr.exe', 'eventvwr.exe'),
                'taskschd': ('Планировщик заданий', 'miguiresource.dll,1', 'Control schedtasks'),
                'useracc': ('Учетные записи', 'imageres.dll,74', 'Control userpasswords'),
                'useracc2': ('Учетные записи (классич.)', 'shell32.dll,111', 'Control userpasswords2'),
                'network': ('Сетевые подключения', 'ncpa.cpl', 'control ncpa.cpl'),
                'wf': ('Брандмауэр Windows', 'wscui.cpl,3', 'wf.msc'),
                'run': ('Выполнить', 'shell32.dll,24',
                        'explorer shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'),
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
                           'explorer shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}')
            }

            for cmd_name, (display_name, icon, command) in entries.items():
                try:
                    cmd_key_path = f"{command_store_path}\\{cmd_name}"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, cmd_key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, cmd_key_path, 0,
                                        winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, display_name)
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon)

                    cmd_command_path = f"{cmd_key_path}\\command"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, cmd_command_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, cmd_command_path, 0,
                                        winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
                except Exception:
                    continue
            return True
        except Exception:
            return False

    # ---------- Администрирование / Система ----------

    def add_admin_menu(self):
        try:
            self.create_command_store_entries()
            clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu)
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Администрирование")
                winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ,
                                  "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run")
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "mmc.exe")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Top")
            self.update_menu_items_status()
        except Exception:
            if not self.run_as_admin_add_admin_menu():
                messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Администрирование'")

    def run_as_admin_add_admin_menu(self):
        try:
            ps_command = '''
            $commandStorePath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell"
            $adminCommands = @{
                'cmd' = @('Командная строка', 'cmd.exe', 'cmd.exe')
                'perfmon' = @('Монитор ресурсов', 'imageres.dll,144', 'perfmon.exe /res')
                'relmon' = @('Монитор надежности системы', 'perfmon.exe', 'perfmon.exe /rel')
                'trouble' = @('Устранение неполадок', 'imageres.dll,124', 'control /name Microsoft.Troubleshooting')
                'services' = @('Службы', 'filemgmt.dll', 'mmc.exe services.msc')
                'event' = @('Просмотр событий', 'eventvwr.exe', 'eventvwr.exe')
                'taskschd' = @('Планировщик заданий', 'miguiresource.dll,1', 'Control schedtasks')
                'useracc' = @('Учетные записи', 'imageres.dll,74', 'Control userpasswords')
                'useracc2' = @('Учетные записи (классич.)', 'shell32.dll,111', 'Control userpasswords2')
                'network' = @('Сетевые подключения', 'ncpa.cpl', 'control ncpa.cpl')
                'wf' = @('Брандмауэр Windows', 'wscui.cpl,3', 'wf.msc')
                'run' = @('Выполнить', 'shell32.dll,24', 'explorer shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}')
            }
            foreach ($cmdName in $adminCommands.Keys) {
                $displayName, $icon, $command = $adminCommands[$cmdName]
                $cmdKeyPath = "$commandStorePath\\$cmdName"
                try {
                    New-Item -Path $cmdKeyPath -Force | Out-Null
                    Set-ItemProperty -Path $cmdKeyPath -Name "(default)" -Value $displayName -Type String
                    Set-ItemProperty -Path $cmdKeyPath -Name "Icon" -Value $icon -Type String
                    $cmdCommandPath = "$cmdKeyPath\\command"
                    New-Item -Path $cmdCommandPath -Force | Out-Null
                    Set-ItemProperty -Path $cmdCommandPath -Name "(default)" -Value $command -Type String
                } catch { }
            }
            $clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"
            try {
                New-Item -Path $clsidPathHKCU -Force -ErrorAction Stop | Out-Null
                Set-ItemProperty -Path $clsidPathHKCU -Name "MUIVerb" -Value "Администрирование" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "SubCommands" -Value "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "Icon" -Value "mmc.exe" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "Position" -Value "Top" -Type String
                return $true
            } catch { return $false }
            '''
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(["powershell", "-Command", ps_command],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=30, startupinfo=startupinfo, shell=True)
            self.update_menu_items_status()
            return result.returncode == 0
        except Exception:
            return False

    def remove_admin_menu(self):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin')
        except Exception:
            pass
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE,
                             r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin')
        except Exception:
            pass
        self.update_menu_items_status()
        return True

    def add_system_menu(self):
        try:
            self.create_command_store_entries()
            clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'
            winreg.CreateKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu)
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Система")
                winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ,
                                  "control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;defrag;folderoptions;networksh;appwiz;power;rstrui;search")
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "imageres.dll,104")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Top")
            self.update_menu_items_status()
        except Exception:
            if not self.run_as_admin_add_system_menu():
                messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Система'")

    def run_as_admin_add_system_menu(self):
        try:
            ps_command = '''
            $commandStorePath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell"
            $systemCommands = @{
                'control' = @('Панель управления', 'control.exe', 'control.exe')
                'sysdm' = @('Свойства системы', 'sysdm.cpl', 'control sysdm.cpl')
                'msconfig' = @('Конфигурация системы', 'msconfig.exe', 'msconfig.exe')
                'regedit' = @('Редактор реестра', 'regedit.exe', 'regedit.exe')
                'taskmgr' = @('Диспетчер задач', 'taskmgr.exe', 'taskmgr.exe')
                'devmgr' = @('Диспетчер устройств', 'devmgr.dll,4', 'mmc.exe devmgmt.msc')
                'cleanmgr' = @('Очистка диска', 'cleanmgr.exe', 'cleanmgr.exe')
                'defrag' = @('Дефрагментация диска', 'dfrgui.exe', 'dfrgui.exe')
                'folderoptions' = @('Свойства папки', 'explorer.exe', 'control folders')
                'networksh' = @('Центр управления сетями', 'networkexplorer.dll,4', 'control /name Microsoft.NetworkAndSharingCenter')
                'appwiz' = @('Программы и компоненты', 'appwiz.cpl', 'control appwiz.cpl')
                'power' = @('Электропитание', 'powercfg.cpl', 'control powercfg.cpl')
                'rstrui' = @('Восстановление системы', 'rstrui.exe', 'rstrui.exe')
                'search' = @('Поиск', 'shell32.dll,22', 'explorer shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}')
            }
            foreach ($cmdName in $systemCommands.Keys) {
                $displayName, $icon, $command = $systemCommands[$cmdName]
                $cmdKeyPath = "$commandStorePath\\$cmdName"
                try {
                    New-Item -Path $cmdKeyPath -Force | Out-Null
                    Set-ItemProperty -Path $cmdKeyPath -Name "(default)" -Value $displayName -Type String
                    Set-ItemProperty -Path $cmdKeyPath -Name "Icon" -Value $icon -Type String
                    $cmdCommandPath = "$cmdKeyPath\\command"
                    New-Item -Path $cmdCommandPath -Force | Out-Null
                    Set-ItemProperty -Path $cmdCommandPath -Name "(default)" -Value $command -Type String
                } catch { }
            }
            $clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"
            try {
                New-Item -Path $clsidPathHKCU -Force -ErrorAction Stop | Out-Null
                Set-ItemProperty -Path $clsidPathHKCU -Name "MUIVerb" -Value "Система" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "SubCommands" -Value "control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;defrag;folderoptions;networksh;appwiz;power;rstrui;search" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "Icon" -Value "imageres.dll,104" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "Position" -Value "Top" -Type String
                return $true
            } catch { return $false }
            '''
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(["powershell", "-Command", ps_command],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=30, startupinfo=startupinfo, shell=True)
            self.update_menu_items_status()
            return result.returncode == 0
        except Exception:
            return False

    def remove_system_menu(self):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System')
        except Exception:
            pass
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE,
                             r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System')
        except Exception:
            pass
        self.update_menu_items_status()
        return True

    # ---------- Меню рабочего стола ----------

    def add_regedit(self):
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\Regedit', "Редактор реестра",
            "regedit.exe,0", "regedit")

    def remove_regedit(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\Regedit')

    def add_restart_explorer(self):
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\RestartExplorer', "Перезапустить explorer",
            "explorer.exe", "cmd.exe /c taskkill /f /im explorer.exe & start explorer.exe")

    def remove_restart_explorer(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\RestartExplorer')

    def add_task_manager(self):
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\TaskManager', "Диспетчер задач",
            "taskmgr.exe,0", "taskmgr.exe")

    def remove_task_manager(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\TaskManager')

    def add_clipboard_clear(self):
        imaster_dll_path = self.check_imaster_dll_from_main()
        self._create_desktop_shell_item(
            r'DesktopBackground\Shell\Clip', "Очистить буфер",
            f"{imaster_dll_path},-1003" if imaster_dll_path else None,
            "cmd /c echo off | clip")
        self.update_power_menu_icons()

    def remove_clipboard_clear(self):
        self._delete_desktop_shell_item(r'DesktopBackground\Shell\Clip')

    def _create_desktop_shell_item(self, key_path, mui_verb, icon, command):
        try:
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, mui_verb)
                winreg.SetValueEx(key, "Extended", 0, winreg.REG_SZ, "")
                if icon:
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon)

            command_key = key_path + r'\command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)

            self.update_desktop_shell_order()
            self.update_menu_items_status()
        except PermissionError:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        self._create_desktop_shell_item(key_path, mui_verb, icon, command)
                    except Exception:
                        messagebox.showerror("Ошибка", f"Не удалось добавить пункт '{mui_verb}'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def _delete_desktop_shell_item(self, key_path):
        try:
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path + r'\command')
            except FileNotFoundError:
                pass
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
            self.update_desktop_shell_order()
            self.update_menu_items_status()
        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except Exception:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def check_imaster_dll_from_main(self):
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
        imaster_dll_path = self.check_imaster_dll_from_main()
        if not imaster_dll_path:
            for key_path in [r'DesktopBackground\Shell\Reboot',
                             r'DesktopBackground\Shell\Hibernation',
                             r'DesktopBackground\Shell\Clip']:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.DeleteValue(key, "Icon")
                except Exception:
                    pass
            return

        for key_path, icon in [
            (r'DesktopBackground\Shell\Reboot', f"{imaster_dll_path},-1001"),
            (r'DesktopBackground\Shell\Hibernation', f"{imaster_dll_path},-1002"),
            (r'DesktopBackground\Shell\Clip', f"{imaster_dll_path},-1003"),
        ]:
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, icon)
            except Exception:
                pass

    def add_power_menu(self):
        try:
            imaster_dll_path = self.check_imaster_dll_from_main()

            hibernate_key = r'DesktopBackground\Shell\Hibernation'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, hibernate_key)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, hibernate_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Гибернация")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "Extended", 0, winreg.REG_SZ, "")
                if imaster_dll_path:
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, f"{imaster_dll_path},-1002")

            hibernate_command = hibernate_key + r'\command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, hibernate_command)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, hibernate_command, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "shutdown /h")

            reboot_key = r'DesktopBackground\Shell\Reboot'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, reboot_key)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reboot_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Перезагрузка компьютера")
                if imaster_dll_path:
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, f"{imaster_dll_path},-1001")

            reboot_command = reboot_key + r'\Command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, reboot_command)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reboot_command, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "shutdown /r /t 0")

            shutdown_key = r'DesktopBackground\Shell\Shutdown'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shutdown_key)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shutdown_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ,
                                  "%SystemRoot%\\System32\\shell32.dll,-28")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Выключение компьютера")

            shutdown_command = shutdown_key + r'\Command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shutdown_command)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shutdown_command, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "shutdown /s /t 0")

            self.update_desktop_shell_order()
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении меню питания: {str(e)}")

    def remove_power_menu(self):
        try:
            for key_path in [
                r'DesktopBackground\Shell\Hibernation\command',
                r'DesktopBackground\Shell\Hibernation',
                r'DesktopBackground\Shell\Reboot\Command',
                r'DesktopBackground\Shell\Reboot',
                r'DesktopBackground\Shell\Shutdown\Command',
                r'DesktopBackground\Shell\Shutdown'
            ]:
                try:
                    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
                except Exception:
                    pass
            self.update_desktop_shell_order()
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении меню питания: {str(e)}")

    # ---------- Проверка существования ----------

    def check_menu_item_exists(self, item_type):
        try:
            if item_type in ("admin_menu", "system_menu"):
                suffix = "Admin" if item_type == "admin_menu" else "System"
                for hive, base in [
                    (winreg.HKEY_CURRENT_USER, 'Software\\Classes\\CLSID'),
                    (winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Classes\\CLSID')
                ]:
                    try:
                        path = f'{base}\\{{20D04FE0-3AEA-1069-A2D8-08002B30309D}}\\shell\\{suffix}'
                        with winreg.OpenKey(hive, path) as key:
                            winreg.QueryValueEx(key, "MUIVerb")
                            return True
                    except FileNotFoundError:
                        continue
                return False

            desktop_keys = {
                "regedit": (r'DesktopBackground\Shell\Regedit', "Редактор реестра"),
                "restart_explorer": (r'DesktopBackground\Shell\RestartExplorer', "Перезапустить explorer"),
                "clipboard_clear": (r'DesktopBackground\Shell\Clip', "Очистить буфер"),
                "task_manager": (r'DesktopBackground\Shell\TaskManager', "Диспетчер задач"),
            }
            if item_type in desktop_keys:
                path, verb = desktop_keys[item_type]
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, path) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == verb
                except FileNotFoundError:
                    return False

            if item_type == "power_menu":
                for key_path in [r'DesktopBackground\Shell\Hibernation',
                                 r'DesktopBackground\Shell\Reboot',
                                 r'DesktopBackground\Shell\Shutdown']:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                            return True
                    except FileNotFoundError:
                        continue
                return False

            if item_type in ("copy_as_path", "pin_to_start", "pin_to_taskbar"):
                paths = {
                    "copy_as_path": (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu',
                                     "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}"),
                    "pin_to_start": (r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen',
                                     "{470C0EBD-5D73-4d58-9CED-E91E22E23282}"),
                    "pin_to_taskbar": (r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}',
                                       "Taskband Pin"),
                }
                path, expected = paths[item_type]
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                        value, _ = winreg.QueryValueEx(key, "")
                        return value == expected
                except FileNotFoundError:
                    return False

            blocked_map = {
                "send_to_device": "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}",
                "compatibility": "{1d27f844-3a1f-4410-85ac-14651078412d}",
                "previous_version": "{596AB062-B4D2-4215-9F74-E9109B0A8153}",
                "share_access": "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}",
            }
            if item_type in blocked_map:
                blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path) as key:
                        try:
                            winreg.QueryValueEx(key, blocked_map[item_type])
                            return False
                        except FileNotFoundError:
                            return True
                except FileNotFoundError:
                    return True

            if item_type == "modern_sharing":
                paths = [
                    r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
                    r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
                    r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
                    r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
                    r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing'
                ]
                for path in paths:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                            value, _ = winreg.QueryValueEx(key, "")
                            if value == "{e2bf9676-5f8f-435c-97eb-11607a5bedf7}":
                                return True
                    except Exception:
                        continue
                return False

            if item_type == "defender":
                for path in [r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}',
                             r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no']:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                            return path.endswith("A780}")
                    except FileNotFoundError:
                        continue
                return False

            if item_type == "paint3d":
                for ext in ['bmp', 'fbx', 'gif']:
                    try:
                        with winreg.OpenKey(
                                winreg.HKEY_LOCAL_MACHINE,
                                f"SOFTWARE\\Classes\\SystemFileAssociations\\.{ext}\\Shell\\3D Edit"):
                            return True
                    except Exception:
                        continue
                return False

            if item_type == "delete_folder_content":
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                        r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'):
                        return True
                except FileNotFoundError:
                    return False

            if item_type == "copy_move":
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                        r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To'):
                        return True
                except FileNotFoundError:
                    return False

            if item_type == "take_ownership":
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'*\shell\runas'):
                        return True
                except FileNotFoundError:
                    return False
        except Exception:
            return False
        return False

    # ---------- Остальные добавления/удаления ----------

    def add_copy_as_path(self):
        key_path = r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu'
        self._simple_lm_value(key_path, "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}")

    def remove_copy_as_path(self):
        self._simple_lm_delete(r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu')

    def add_pin_to_start(self):
        key_path = r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen'
        self._simple_lm_value(key_path, "{470C0EBD-5D73-4d58-9CED-E91E22E23282}")

    def remove_pin_to_start(self):
        self._simple_lm_delete(r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen')

    def add_pin_to_taskbar(self):
        key_path = r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}'
        self._simple_lm_value(key_path, "Taskband Pin")

    def remove_pin_to_taskbar(self):
        self._simple_lm_delete(
            r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}')

    def _simple_lm_value(self, key_path, value):
        try:
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, value)
            self.update_menu_items_status()
        except PermissionError:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path) and \
                    self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, value)
                    self.update_menu_items_status()
                except Exception:
                    messagebox.showerror("Ошибка", "Не удалось добавить пункт")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить права")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def _simple_lm_delete(self, key_path):
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            self.update_menu_items_status()
        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path) and \
                    self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                try:
                    self.delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE, key_path)
                    self.update_menu_items_status()
                except Exception:
                    messagebox.showerror("Ошибка", "Не удалось удалить пункт")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить права")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def delete_registry_key_recursive(self, hive, key_path):
        try:
            with winreg.OpenKey(hive, key_path) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        self.delete_registry_key_recursive(hive, key_path + "\\" + subkey_name)
                    except OSError:
                        break
                    i += 1
        except FileNotFoundError:
            return
        except Exception:
            pass
        try:
            winreg.DeleteKey(hive, key_path)
        except Exception:
            pass

    def add_send_to_device(self):
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        try:
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}")
                except FileNotFoundError:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_send_to_device(self):
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        try:
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}", 0,
                                  winreg.REG_SZ, "Передать на устройство")
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_paint3d_edit(self):
        try:
            for ext in ['bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                        'obj', 'ply', 'png', 'stl', 'tif', 'tiff']:
                try:
                    shell_key_path = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{ext}\\Shell"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, shell_key_path)
                    paint3d_key = f"{shell_key_path}\\3D Edit"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                                          "@%SystemRoot%\\system32\\mspaint.exe,-59500")
                    command_key = f"{paint3d_key}\\command"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_key, 0,
                                        winreg.KEY_SET_VALUE) as cmd_key:
                        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_EXPAND_SZ,
                                          '%SystemRoot%\\system32\\mspaint.exe "%1" /ForceBootstrapPaint3D')
                except Exception:
                    continue
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_paint3d_edit(self):
        try:
            for ext in ['bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                        'obj', 'ply', 'png', 'stl', 'tif', 'tiff']:
                paint3d_key = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{ext}\\Shell\\3D Edit"
                try:
                    winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key + "\\command")
                except FileNotFoundError:
                    pass
                except Exception:
                    pass
                try:
                    winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key)
                except FileNotFoundError:
                    pass
                except Exception:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_defender_scan(self):
        try:
            defender_clsid_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
            defender_clsid_no_path = defender_clsid_path + '_no'

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_no_path):
                    if self.rename_registry_key(winreg.HKEY_LOCAL_MACHINE,
                                                defender_clsid_no_path, defender_clsid_path):
                        self.update_menu_items_status()
                        return
            except Exception:
                pass

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path):
                    self.update_menu_items_status()
                    return
            except Exception:
                pass

            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Microsoft.Windows.Defender.IconHandler")

            inproc_path = defender_clsid_path + r'\InprocServer32'
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, inproc_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, inproc_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_EXPAND_SZ,
                                  "%SystemRoot%\\System32\\WindowsDefender\\ShellExt.dll")
                winreg.SetValueEx(key, "ThreadingModel", 0, winreg.REG_SZ, "Apartment")

            self.update_menu_items_status()
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при добавлении пункта")

    def remove_defender_scan(self):
        try:
            defender_clsid_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
            defender_clsid_no_path = defender_clsid_path + '_no'

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path):
                    if self.rename_registry_key(winreg.HKEY_LOCAL_MACHINE,
                                                defender_clsid_path, defender_clsid_no_path):
                        self.update_menu_items_status()
                        return True
                    messagebox.showerror("Ошибка", "Не удалось удалить пункт")
                    return False
            except FileNotFoundError:
                self.update_menu_items_status()
                return True
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при удалении пункта")
            return False

    def add_modern_sharing(self):
        try:
            clsid = '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'
            for key_path in [
                r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing'
            ]:
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0,
                                        winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, clsid)
                except Exception:
                    continue
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_modern_sharing(self):
        try:
            for key_path in [
                r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing'
            ]:
                try:
                    winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                except Exception:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_delete_folder_content(self):
        try:
            key_path = r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Удалить содержимое папки")
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "shell32.dll,-240")

            command_key = key_path + r'\command'
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_key, 0, winreg.KEY_SET_VALUE) as cmd_key:
                winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ,
                                  'cmd.exe /c "cd /d \"%1\" && del /q /f /s *.*"')
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_delete_folder_content(self):
        try:
            for key_path in [r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent\command',
                             r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent']:
                try:
                    winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                except Exception:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_copy_move_menu(self):
        try:
            for key_path, clsid in [
                (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
                 "{C2FBB630-2971-11D1-A18C-00C04FD75D13}"),
                (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To',
                 "{C2FBB631-2971-11D1-A18C-00C04FD75D13}")
            ]:
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0,
                                        winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, clsid)
                except Exception:
                    continue
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_copy_move_menu(self):
        try:
            for key_path in [r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
                             r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To']:
                try:
                    winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                except Exception:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_take_ownership(self):
        try:
            for base_key, cmd in [
                (r'*\shell\runas',
                 'cmd.exe /c takeown /f "%1" && icacls "%1" /grant администраторы:F'),
                (r'Directory\shell\runas',
                 'cmd.exe /c takeown /f "%1" /r /d y && icacls "%1" /grant администраторы:F /t')
            ]:
                winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, base_key)
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, base_key, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Стать владельцем")
                    winreg.SetValueEx(key, "NoWorkingDirectory", 0, winreg.REG_SZ, "")
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "imageres.dll,54")
                    winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Стать владельцем")

                command_key = base_key + r'\command'
                winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0,
                                    winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, cmd)
                    winreg.SetValueEx(key, "IsolatedCommand", 0, winreg.REG_SZ, cmd)
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_take_ownership(self):
        try:
            for key_path in [r'*\shell\runas\command', r'*\shell\runas',
                             r'Directory\shell\runas\command', r'Directory\shell\runas']:
                try:
                    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
                except Exception:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_compatibility(self):
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{1d27f844-3a1f-4410-85ac-14651078412d}")
                except FileNotFoundError:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_compatibility(self):
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "{1d27f844-3a1f-4410-85ac-14651078412d}", 0, winreg.REG_SZ,
                                  "Исправление проблем с совместимостью")
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_previous_version(self):
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{596AB062-B4D2-4215-9F74-E9109B0A8153}")
                except FileNotFoundError:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_previous_version(self):
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "{596AB062-B4D2-4215-9F74-E9109B0A8153}", 0, winreg.REG_SZ,
                                  "Восстановить прежнюю версию")
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def add_share_access(self):
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}")
                except FileNotFoundError:
                    pass
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def remove_share_access(self):
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}", 0, winreg.REG_SZ, "")
            self.update_menu_items_status()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    # ---------- UI-компоненты ----------

    def create_menu_item_button(self, parent, item):
        btn_frame = ttk.Frame(parent)
        btn_frame.pack()

        indicator_canvas = tk.Canvas(btn_frame, width=20, height=20, highlightthickness=0)
        indicator_canvas.pack(side=tk.LEFT, padx=(0, 5))
        indicator_canvas.create_oval(2, 2, 18, 18, outline='#e74c3c', fill='#e74c3c')
        indicator_canvas.create_text(10, 10, text="✗", fill="white", font=('Arial', 10, 'bold'))

        item["indicator_canvas"] = indicator_canvas
        item["indicator_circle"] = indicator_canvas.find_all()[0]
        item["indicator_text"] = indicator_canvas.find_all()[1]

        btn = ttk.Button(btn_frame, text=item["name"], width=50,
                         command=lambda i=item: self.toggle_menu_item(i))
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=2)
        return btn

    def update_indicator(self, item, is_enabled):
        if 'indicator_canvas' in item and item['indicator_canvas']:
            canvas = item['indicator_canvas']
            color = '#27ae60' if is_enabled else '#e74c3c'
            text = "✓" if is_enabled else "✗"
            canvas.itemconfig(item['indicator_circle'], outline=color, fill=color)
            canvas.itemconfig(item['indicator_text'], text=text)

    def restart_explorer(self):
        try:
            subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True)
            subprocess.Popen("explorer.exe")
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при перезагрузке проводника")

    def set_default_settings(self):
        try:
            for item in self.menu_item_buttons:
                item_data = item["data"]
                current_state = item_data["check_command"]()
                if item_data.get("default_enabled", True) and not current_state:
                    try:
                        item_data["add_command"]()
                    except Exception:
                        pass
                elif not item_data.get("default_enabled", True) and current_state:
                    try:
                        item_data["remove_command"]()
                    except Exception:
                        pass
            self.update_power_menu_icons()
            self.update_menu_items_status()
            self.update_desktop_shell_order()
            messagebox.showinfo("Успех", "Настройки сброшены по умолчанию!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сбросе настроек: {str(e)}")

    def toggle_menu_item(self, item):
        try:
            exists = item["check_command"]()
            if exists:
                item["remove_command"]()
            else:
                item["add_command"]()
            self.update_indicator(item, not exists)
            self.update_menu_items_status()
            if item["name"] in ["Выключение/Перезагрузка/Гибернация", "Диспетчер задач",
                                "Перезапустить explorer", "Очистить буфер", "Редактор реестра"]:
                self.update_desktop_shell_order()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта: {str(e)}")

    def update_menu_items_status(self):
        try:
            for btn_info in self.menu_item_buttons:
                item = btn_info["data"]
                exists = item["check_command"]()
                self.update_indicator(item, exists)
                btn_info["button"].configure(text=item['name'])
        except Exception:
            pass

    def set_registry_key_permissions_simple(self, key_path):
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(["reg", "add", f"HKLM\\{key_path}", "/f"],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=10, startupinfo=startupinfo, shell=True)
            if result.returncode == 0:
                return True

            ps_command = f'''
            try {{
                New-Item -Path "HKLM:\\{key_path}" -Force -ErrorAction Stop | Out-Null
                return $true
            }} catch {{ return $false }}
            '''
            result = subprocess.run(["powershell", "-Command", ps_command],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=10, startupinfo=startupinfo, shell=True)
            return result.returncode == 0
        except Exception:
            return False

    def take_ownership_registry_key_simple(self, key_path):
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(["reg", "add", f"HKLM\\{key_path}", "/f", "/reg:64"],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=10, startupinfo=startupinfo, shell=True)
            if result.returncode == 0:
                return True

            ps_command = f'''
            try {{
                $key = "HKLM:\\{key_path}"
                $acl = Get-Acl $key -ErrorAction Stop
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("BUILTIN\\Administrators","FullControl","Allow")
                $acl.SetAccessRule($rule)
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("NT AUTHORITY\\SYSTEM","FullControl","Allow")
                $acl.SetAccessRule($rule)
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("{os.environ["USERNAME"]}","FullControl","Allow")
                $acl.SetAccessRule($rule)
                Set-Acl $key $acl -ErrorAction Stop
                return $true
            }} catch {{
                try {{
                    New-Item -Path "HKLM:\\{key_path}" -Force -ErrorAction Stop | Out-Null
                    return $true
                }} catch {{ return $false }}
            }}
            '''
            result = subprocess.run(["powershell", "-Command", ps_command],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=10, startupinfo=startupinfo, shell=True)
            return result.returncode == 0
        except Exception:
            return False

    def set_registry_key_permissions(self, hive, key_path):
        try:
            root = {winreg.HKEY_CLASSES_ROOT: "HKCR", winreg.HKEY_CURRENT_USER: "HKCU",
                    winreg.HKEY_LOCAL_MACHINE: "HKLM"}.get(hive, "HKCR")
            full_path = f"{root}\\{key_path}"

            ps_command = f'''
            $key = "Registry::{full_path}"
            try {{
                $acl = Get-Acl $key
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("BUILTIN\\Administrators","FullControl","Allow")
                $acl.SetAccessRule($rule)
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("NT AUTHORITY\\SYSTEM","FullControl","Allow")
                $acl.SetAccessRule($rule)
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("{os.environ["USERNAME"]}","FullControl","Allow")
                $acl.SetAccessRule($rule)
                Set-Acl $key $acl
            }} catch {{ }}
            '''
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(["powershell", "-Command", ps_command],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=30, startupinfo=startupinfo)
            return result.returncode == 0
        except Exception:
            return False

    def take_ownership_registry_key(self, hive, key_path):
        try:
            root = {winreg.HKEY_CLASSES_ROOT: "HKCR", winreg.HKEY_CURRENT_USER: "HKCU",
                    winreg.HKEY_LOCAL_MACHINE: "HKLM"}.get(hive, "HKCR")
            full_path = f"{root}\\{key_path}"
            ps_command = f'''
            $key = "Registry::{full_path}"
            try {{
                $acl = Get-Acl $key
                $owner = [System.Security.Principal.NTAccount]"BUILTIN\\Administrators"
                $acl.SetOwner($owner)
                Set-Acl $key $acl
            }} catch {{ }}
            '''
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(["powershell", "-Command", ps_command],
                                    capture_output=True, text=True, encoding='cp866',
                                    timeout=30, startupinfo=startupinfo)
            return result.returncode == 0
        except Exception:
            return False

    def rename_registry_key(self, hive, old_key_path, new_key_path):
        try:
            self.copy_registry_key(hive, old_key_path, new_key_path)
            self.delete_registry_key(hive, old_key_path)
            return True
        except Exception:
            return False

    def copy_registry_key(self, hive, source_path, dest_path):
        winreg.CreateKey(hive, dest_path)
        with winreg.OpenKey(hive, source_path) as source_key:
            with winreg.OpenKey(hive, dest_path, 0, winreg.KEY_WRITE) as dest_key:
                i = 0
                while True:
                    try:
                        name, value, type_val = winreg.EnumValue(source_key, i)
                        winreg.SetValueEx(dest_key, name, 0, type_val, value)
                        i += 1
                    except OSError:
                        break
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(source_key, i)
                        self.copy_registry_key(hive, source_path + "\\" + subkey_name,
                                               dest_path + "\\" + subkey_name)
                        i += 1
                    except OSError:
                        break
        return True

    def delete_registry_key(self, hive, key_path):
        try:
            with winreg.OpenKey(hive, key_path) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        self.delete_registry_key(hive, key_path + "\\" + subkey_name)
                        i += 1
                    except OSError:
                        break
        except FileNotFoundError:
            return True
        winreg.DeleteKey(hive, key_path)
        return True