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
        self.powershell_log_window = None

        self.create_widgets()
        self.show()

    def create_widgets(self):
        """Создание виджетов для настройки контекстного меню"""

        title_label = ttk.Label(
            self.frame,
            text="Настройка контекстного меню",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        notebook = ttk.Notebook(self.frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        desktop_menu_frame = ttk.Frame(notebook)
        create_menu_frame = ttk.Frame(notebook)
        sendto_menu_frame = ttk.Frame(notebook)

        notebook.add(desktop_menu_frame, text="     Меню рабочего стола     ")
        notebook.add(create_menu_frame, text="          Меню Создать          ")
        notebook.add(sendto_menu_frame, text="         Меню Отправить         ")

        self.create_desktop_menu_tab(desktop_menu_frame)
        self.create_new_menu_tab(create_menu_frame)
        self.create_sendto_menu_tab(sendto_menu_frame)

    def create_desktop_menu_tab(self, parent):
        """Создание вкладки меню рабочего стола"""

        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        title_label = ttk.Label(
            scrollable_frame,
            text="Управление пунктами контекстного меню",
            font=('Arial', 12, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        control_frame = ttk.Frame(scrollable_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=10)

        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)

        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        self.restart_explorer_btn = ttk.Button(
            buttons_frame,
            text="Перезапустить Проводник",
            command=self.restart_explorer,
            width=25
        )
        self.restart_explorer_btn.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Button(
            buttons_frame,
            text="По умолчанию",
            command=self.set_default_settings,
            width=15
        ).pack(side=tk.LEFT, padx=5, pady=5)

        self.menu_items_data = [
            {
                "name": "Изменить с помощью Paint 3D",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_paint3d_edit,
                "remove_command": self.remove_paint3d_edit,
                "check_command": lambda: self.check_menu_item_exists("paint3d"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Исправление проблем с совместимостью",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_compatibility,
                "remove_command": self.remove_compatibility,
                "check_command": lambda: self.check_menu_item_exists("compatibility"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Закрепить на начальном экране",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_pin_to_start,
                "remove_command": self.remove_pin_to_start,
                "check_command": lambda: self.check_menu_item_exists("pin_to_start"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Закрепить на Панели задач",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_pin_to_taskbar,
                "remove_command": self.remove_pin_to_taskbar,
                "check_command": lambda: self.check_menu_item_exists("pin_to_taskbar"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Копировать как путь",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_copy_as_path,
                "remove_command": self.remove_copy_as_path,
                "check_command": lambda: self.check_menu_item_exists("copy_as_path"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Отправить (metro)",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_modern_sharing,
                "remove_command": self.remove_modern_sharing,
                "check_command": lambda: self.check_menu_item_exists("modern_sharing"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Передать на устройство",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_send_to_device,
                "remove_command": self.remove_send_to_device,
                "check_command": lambda: self.check_menu_item_exists("send_to_device"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Предоставить доступ к",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_share_access,
                "remove_command": self.remove_share_access,
                "check_command": lambda: self.check_menu_item_exists("share_access"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Проверка с использованием Windows Defender",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_defender_scan,
                "remove_command": self.remove_defender_scan,
                "check_command": lambda: self.check_menu_item_exists("defender"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Восстановить прежнюю версию",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_previous_version,
                "remove_command": self.remove_previous_version,
                "check_command": lambda: self.check_menu_item_exists("previous_version"),
                "enabled": False,
                "default_enabled": True
            },
            {
                "name": "Стать владельцем",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_take_ownership,
                "remove_command": self.remove_take_ownership,
                "check_command": lambda: self.check_menu_item_exists("take_ownership"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Копировать/Переместить в папку",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_copy_move_menu,
                "remove_command": self.remove_copy_move_menu,
                "check_command": lambda: self.check_menu_item_exists("copy_move"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Удалить содержимое папки",
                "add_tooltip": "",
                "remove_tooltip": "",
                "add_command": self.add_delete_folder_content,
                "remove_command": self.remove_delete_folder_content,
                "check_command": lambda: self.check_menu_item_exists("delete_folder_content"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Диспетчер задач",
                "add_tooltip": "Контекстное меню Рабочего стола",
                "remove_tooltip": "Контекстное меню Рабочего стола",
                "add_command": self.add_task_manager,
                "remove_command": self.remove_task_manager,
                "check_command": lambda: self.check_menu_item_exists("task_manager"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Перезапустить explorer",
                "add_tooltip": "Контекстное меню Рабочего стола",
                "remove_tooltip": "Контекстное меню Рабочего стола",
                "add_command": self.add_restart_explorer,
                "remove_command": self.remove_restart_explorer,
                "check_command": lambda: self.check_menu_item_exists("restart_explorer"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Редактор реестра",
                "add_tooltip": "Контекстное меню Рабочего стола",
                "remove_tooltip": "Контекстное меню Рабочего стола",
                "add_command": self.add_regedit,
                "remove_command": self.remove_regedit,
                "check_command": lambda: self.check_menu_item_exists("regedit"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Выключение/Перезагрузка/Гибернация",
                "add_tooltip": "Контекстное меню Рабочего стола",
                "remove_tooltip": "Контекстное меню Рабочего стола",
                "add_command": self.add_power_menu,
                "remove_command": self.remove_power_menu,
                "check_command": lambda: self.check_menu_item_exists("power_menu"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Администрирование",
                "add_tooltip": "Контекстное меню Компьютер",
                "remove_tooltip": "Контекстное меню Компьютер",
                "add_command": self.add_admin_menu,
                "remove_command": self.remove_admin_menu,
                "check_command": lambda: self.check_menu_item_exists("admin_menu"),
                "enabled": False,
                "default_enabled": False
            },
            {
                "name": "Система",
                "add_tooltip": "Контекстное меню Компьютер",
                "remove_tooltip": "Контекстное меню Компьютер",
                "add_command": self.add_system_menu,
                "remove_command": self.remove_system_menu,
                "check_command": lambda: self.check_menu_item_exists("system_menu"),
                "enabled": False,
                "default_enabled": False
            }
        ]

        self.menu_item_buttons = []
        for item in self.menu_items_data:
            button_frame = ttk.Frame(scrollable_frame)
            button_frame.pack(fill=tk.X, padx=10, pady=5)

            btn = self.create_menu_item_button(button_frame, item)
            self.menu_item_buttons.append({"button": btn, "data": item})

            if item["add_tooltip"]:
                self.create_tooltip(btn, item["add_tooltip"])

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.update_menu_items_status()

    def update_desktop_shell_order(self):
        """Обновление порядка пунктов в меню рабочего стола"""
        try:
            shell_key = r'DesktopBackground\Shell'
            active_items = []

            if self.check_menu_item_exists("task_manager"):
                active_items.append("TaskManager")

            if self.check_menu_item_exists("restart_explorer"):
                active_items.append("RestartExplorer")

            if self.check_menu_item_exists("regedit"):
                active_items.append("Regedit")

            if self.check_menu_item_exists("power_menu"):
                power_keys = [
                    r'DesktopBackground\Shell\Shutdown',
                    r'DesktopBackground\Shell\Reboot',
                    r'DesktopBackground\Shell\Hibernation'
                ]

                for key_path in power_keys:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                            if "Shutdown" in key_path:
                                active_items.append("Shutdown")
                            elif "Reboot" in key_path:
                                active_items.append("Reboot")
                            elif "Hibernation" in key_path:
                                active_items.append("Hibernation")
                    except FileNotFoundError:
                        pass

            active_items.append("Display")
            active_items.append("Personalize")

            order_string = ",".join(active_items)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shell_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, order_string)

        except:
            try:
                winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shell_key)
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shell_key, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Display,Personalize")
            except:
                pass

    def create_command_store_entries(self):
        """Создание всех необходимых записей в CommandStore"""
        try:
            command_store_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell'

            command_store_entries = {
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
                'run': ('Выполнить', 'shell32.dll,24', 'explorer shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'),
                'control': ('Панель управления', 'control.exe', 'control.exe'),
                'sysdm': ('Свойства системы', 'sysdm.cpl', 'control sysdm.cpl'),
                'msconfig': ('Конфигурация системы', 'msconfig.exe', 'msconfig.exe'),
                'regedit': ('Редактор реестра', 'regedit.exe', 'regedit.exe'),
                'taskmgr': ('Диспетчер задач', 'taskmgr.exe', 'taskmgr.exe'),
                'devmgr': ('Диспетчер устройств', 'devmgr.dll,4', 'mmc.exe devmgmt.msc'),
                'cleanmgr': ('Очистка диска', 'cleanmgr.exe', 'cleanmgr.exe'),
                'defrag': ('Дефрагментация диска', 'dfrgui.exe', 'dfrgui.exe'),
                'folderoptions': ('Свойства папки', 'explorer.exe', 'control folders'),
                'networksh': (
                    'Центр управления сетями', 'networkexplorer.dll,4',
                    'control /name Microsoft.NetworkAndSharingCenter'),
                'appwiz': ('Программы и компоненты', 'appwiz.cpl', 'control appwiz.cpl'),
                'power': ('Электропитание', 'powercfg.cpl', 'control powercfg.cpl'),
                'rstrui': ('Восстановление системы', 'rstrui.exe', 'rstrui.exe'),
                'search': ('Поиск', 'shell32.dll,22', 'explorer shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}')
            }

            for cmd_name, (display_name, icon, command) in command_store_entries.items():
                try:
                    cmd_key_path = f"{command_store_path}\\{cmd_name}"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, cmd_key_path)

                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, cmd_key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, display_name)
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon)

                    cmd_command_path = f"{cmd_key_path}\\command"
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, cmd_command_path)

                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, cmd_command_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)

                except:
                    continue

            return True

        except:
            return False

    def add_admin_menu(self):
        """Добавление пункта 'Администрирование' в контекстное меню Этот компьютер"""
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

        except:
            success = self.run_as_admin_add_admin_menu()
            if not success:
                messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Администрирование'")

    def run_as_admin_add_admin_menu(self):
        """Запуск PowerShell с правами администратора для добавления Администрирования"""
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
                } catch {
                    Write-Host "Ошибка создания команды $cmdName: $_"
                }
            }

            $clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"

            try {
                New-Item -Path $clsidPathHKCU -Force -ErrorAction Stop | Out-Null

                Set-ItemProperty -Path $clsidPathHKCU -Name "MUIVerb" -Value "Администрирование" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "SubCommands" -Value "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "Icon" -Value "mmc.exe" -Type String
                Set-ItemProperty -Path $clsidPathHKCU -Name "Position" -Value "Top" -Type String

                return $true
            } catch {
                $clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"

                try {
                    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
                        Write-Host "Требуются права администратора для записи в HKLM"
                        return $false
                    }

                    New-Item -Path $clsidPathHKLM -Force -ErrorAction Stop | Out-Null

                    Set-ItemProperty -Path $clsidPathHKLM -Name "MUIVerb" -Value "Администрирование" -Type String
                    Set-ItemProperty -Path $clsidPathHKLM -Name "SubCommands" -Value "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run" -Type String
                    Set-ItemProperty -Path $clsidPathHKLM -Name "Icon" -Value "mmc.exe" -Type String
                    Set-ItemProperty -Path $clsidPathHKLM -Name "Position" -Value "Top" -Type String

                    return $true
                } catch {
                    return $false
                }
            }
            '''

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=30,
                startupinfo=startupinfo,
                shell=True
            )

            self.update_menu_items_status()
            return result.returncode == 0

        except:
            return False

    def remove_admin_menu(self):
        """Удаление пункта 'Администрирование' из контекстного меню Этот компьютер"""
        try:
            clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'

            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu)
                removed = True
            except:
                removed = False

            clsid_path_hklm = r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'

            try:
                winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, clsid_path_hklm)
                removed = True
            except:
                pass

            self.update_menu_items_status()
            return True

        except:
            success = self.run_as_admin_remove_admin_menu()
            if not success:
                messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Администрирование'")
            return False

    def run_as_admin_remove_admin_menu(self):
        """Запуск PowerShell с правами администратора для удаления Администрирования"""
        try:
            ps_command = '''
            $clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"
            $clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"

            $removed = $false

            try {
                if (Test-Path $clsidPathHKCU) {
                    Remove-Item -Path $clsidPathHKCU -Recurse -Force -ErrorAction Stop
                    $removed = $true
                }
            } catch { }

            try {
                if (Test-Path $clsidPathHKLM) {
                    Remove-Item -Path $clsidPathHKLM -Recurse -Force -ErrorAction Stop
                    $removed = $true
                }
            } catch { }

            return $removed
            '''

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=30,
                startupinfo=startupinfo,
                shell=True
            )

            self.update_menu_items_status()
            return result.returncode == 0

        except:
            return False

    def add_system_menu(self):
        """Добавление пункта 'Система' в контекстное меню Этот компьютер"""
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

        except:
            success = self.run_as_admin_add_system_menu()
            if not success:
                messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Система'")

    def run_as_admin_add_system_menu(self):
        """Запуск PowerShell с правами администратора для добавления Системы"""
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
            } catch {
                $clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"

                try {
                    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
                        return $false
                    }

                    New-Item -Path $clsidPathHKLM -Force -ErrorAction Stop | Out-Null

                    Set-ItemProperty -Path $clsidPathHKLM -Name "MUIVerb" -Value "Система" -Type String
                    Set-ItemProperty -Path $clsidPathHKLM -Name "SubCommands" -Value "control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;defrag;folderoptions;networksh;appwiz;power;rstrui;search" -Type String
                    Set-ItemProperty -Path $clsidPathHKLM -Name "Icon" -Value "imageres.dll,104" -Type String
                    Set-ItemProperty -Path $clsidPathHKLM -Name "Position" -Value "Top" -Type String

                    return $true
                } catch {
                    return $false
                }
            }
            '''

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=30,
                startupinfo=startupinfo,
                shell=True
            )

            self.update_menu_items_status()
            return result.returncode == 0

        except:
            return False

    def remove_system_menu(self):
        """Удаление пункта 'Система' из контекстного меню Этот компьютер"""
        try:
            clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'

            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu)
                removed = True
            except:
                removed = False

            clsid_path_hklm = r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'

            try:
                winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, clsid_path_hklm)
                removed = True
            except:
                pass

            self.update_menu_items_status()
            return True

        except:
            success = self.run_as_admin_remove_system_menu()
            if not success:
                messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Система'")
            return False

    def run_as_admin_remove_system_menu(self):
        """Запуск PowerShell с правами администратора для удаления Системы"""
        try:
            ps_command = '''
            $clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"
            $clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"

            $removed = $false

            try {
                if (Test-Path $clsidPathHKCU) {
                    Remove-Item -Path $clsidPathHKCU -Recurse -Force -ErrorAction Stop
                    $removed = $true
                }
            } catch { }

            try {
                if (Test-Path $clsidPathHKLM) {
                    Remove-Item -Path $clsidPathHKLM -Recurse -Force -ErrorAction Stop
                    $removed = $true
                }
            } catch { }

            return $removed
            '''

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=30,
                startupinfo=startupinfo,
                shell=True
            )

            self.update_menu_items_status()
            return result.returncode == 0

        except:
            return False

    def add_regedit(self):
        """Добавление пункта 'Редактор реестра' в меню рабочего стола"""
        try:
            key_path = r'DesktopBackground\Shell\Regedit'

            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Редактор реестра")
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "regedit.exe,0")
                winreg.SetValueEx(key, "Extended", 0, winreg.REG_SZ, "")

            command_key = key_path + r'\command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "regedit")

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)

                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Редактор реестра")
                            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "regedit.exe,0")
                            winreg.SetValueEx(key, "Extended", 0, winreg.REG_SZ, "")

                        command_key = key_path + r'\command'
                        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)

                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "regedit")

                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Редактор реестра'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_regedit(self):
        """Удаление пункта 'Редактор реестра' из меню рабочего стола"""
        try:
            key_path = r'DesktopBackground\Shell\Regedit'

            command_key = key_path + r'\command'
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, command_key)
            except FileNotFoundError:
                pass

            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Редактор реестра'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_restart_explorer(self):
        """Добавление пункта 'Перезапустить explorer' в меню рабочего стола"""
        try:
            key_path = r'DesktopBackground\Shell\RestartExplorer'

            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "explorer.exe")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Перезапустить explorer")

            command_key = key_path + r'\command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                                  "cmd.exe /c taskkill /f /im explorer.exe & start explorer.exe")

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)

                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "explorer.exe")
                            winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Перезапустить explorer")

                        command_key = key_path + r'\command'
                        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)

                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                                              "cmd.exe /c taskkill /f /im explorer.exe & start explorer.exe")

                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Перезапустить explorer'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_restart_explorer(self):
        """Удаление пункта 'Перезапустить explorer' из меню рабочего стола"""
        try:
            key_path = r'DesktopBackground\Shell\RestartExplorer'

            command_key = key_path + r'\command'
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, command_key)
            except FileNotFoundError:
                pass

            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Перезапустить explorer'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def check_imaster_dll_from_main(self):
        """Проверка наличия файла imaster.dll через модуль main"""
        try:
            import __main__ as main

            if hasattr(main, 'DLLChecker'):
                dll_info = main.DLLChecker.get_dll_status()

                if dll_info['dll_found'] and dll_info['hash_match']:
                    if 'recommended_path' in dll_info and dll_info['recommended_path']:
                        return dll_info['recommended_path']
                    elif 'dll_path' in dll_info and dll_info['dll_path']:
                        return dll_info['dll_path']
                else:
                    return None
            else:
                return self.check_imaster_dll_direct()

        except:
            return self.check_imaster_dll_direct()

    def check_imaster_dll_direct(self):
        """Прямая проверка наличия файла imaster.dll"""
        try:
            system32_path = os.path.join(os.environ['SYSTEMROOT'], 'System32', 'imaster.dll')
            if os.path.exists(system32_path):
                return system32_path

            if hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            local_imaster_path = os.path.join(base_path, 'imaster.dll')
            if os.path.exists(local_imaster_path):
                return local_imaster_path

            current_dir = os.getcwd()
            work_dir_path = os.path.join(current_dir, 'imaster.dll')
            if os.path.exists(work_dir_path):
                return work_dir_path

            return None

        except:
            return None

    def update_power_menu_icons(self):
        """Обновление иконок для пунктов питания"""
        imaster_dll_path = self.check_imaster_dll_from_main()

        if not imaster_dll_path:
            try:
                reboot_key = r'DesktopBackground\Shell\Reboot'
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reboot_key, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "")
                except:
                    pass

                hibernate_key = r'DesktopBackground\Shell\Hibernation'
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, hibernate_key, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "")
                except:
                    pass
            except:
                pass
            return

        try:
            reboot_key = r'DesktopBackground\Shell\Reboot'
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reboot_key, 0, winreg.KEY_SET_VALUE) as key:
                    if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, f"{imaster_dll_path},-1001")
                    else:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{imaster_dll_path},-1001")
            except:
                pass

            hibernate_key = r'DesktopBackground\Shell\Hibernation'
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, hibernate_key, 0, winreg.KEY_SET_VALUE) as key:
                    if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, f"{imaster_dll_path},-1002")
                    else:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{imaster_dll_path},-1002")
            except:
                pass

        except:
            pass

    def add_power_menu(self):
        """Добавление пунктов выключения, перезагрузки и гибернации"""
        try:
            imaster_dll_path = self.check_imaster_dll_from_main()

            hibernate_key = r'DesktopBackground\Shell\Hibernation'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, hibernate_key)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, hibernate_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Гибернация")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "Extended", 0, winreg.REG_SZ, "")

                if imaster_dll_path:
                    if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, f"{imaster_dll_path},-1002")
                    else:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{imaster_dll_path},-1002")

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
                    if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, f"{imaster_dll_path},-1001")
                    else:
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, f"{imaster_dll_path},-1001")

            reboot_command = reboot_key + r'\Command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, reboot_command)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, reboot_command, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "shutdown /r /t 0")

            shutdown_key = r'DesktopBackground\Shell\Shutdown'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shutdown_key)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shutdown_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_EXPAND_SZ, "%SystemRoot%\\System32\\shell32.dll,-28")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Выключение компьютера")

            shutdown_command = shutdown_key + r'\Command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shutdown_command)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shutdown_command, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "shutdown /s /t 0")

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(r'DesktopBackground\Shell'):
                if self.set_registry_key_permissions_simple(r'DesktopBackground\Shell'):
                    try:
                        self.add_power_menu()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить меню питания")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении меню питания: {str(e)}")

    def remove_power_menu(self):
        """Удаление пунктов выключения, перезагрузки и гибернации"""
        try:
            power_keys = [
                r'DesktopBackground\Shell\Hibernation\command',
                r'DesktopBackground\Shell\Hibernation',
                r'DesktopBackground\Shell\Reboot\Command',
                r'DesktopBackground\Shell\Reboot',
                r'DesktopBackground\Shell\Shutdown\Command',
                r'DesktopBackground\Shell\Shutdown'
            ]

            for key_path in power_keys:
                try:
                    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
                except:
                    pass

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except PermissionError as pe:
            try:
                for key_path in [
                    r'DesktopBackground\Shell\Hibernation',
                    r'DesktopBackground\Shell\Reboot',
                    r'DesktopBackground\Shell\Shutdown'
                ]:
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                    except:
                        pass

                self.update_desktop_shell_order()
                self.update_menu_items_status()
            except:
                messagebox.showerror("Ошибка", "Не удалось удалить меню питания")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении меню питания: {str(e)}")

    def add_task_manager(self):
        """Добавление пункта 'Диспетчер задач' в меню рабочего стола"""
        try:
            key_path = r'DesktopBackground\Shell\TaskManager'

            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "taskmgr.exe,0")
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "")
                winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Диспетчер задач")

            command_key = key_path + r'\Command'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)

            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "taskmgr.exe")

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)

                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "taskmgr.exe,0")
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "")
                            winreg.SetValueEx(key, "Position", 0, winreg.REG_SZ, "Bottom")
                            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, "Диспетчер задач")

                        command_key = key_path + r'\Command'
                        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, command_key)

                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, command_key, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "taskmgr.exe")

                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Диспетчер задач'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_task_manager(self):
        """Удаление пункта 'Диспетчер задач' из меню рабочего стола"""
        try:
            key_path = r'DesktopBackground\Shell\TaskManager'

            command_key = key_path + r'\Command'
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, command_key)
            except FileNotFoundError:
                pass

            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)

            self.update_desktop_shell_order()
            self.update_menu_items_status()

        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError as pe:
            if self.take_ownership_registry_key_simple(key_path):
                if self.set_registry_key_permissions_simple(key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                        self.update_desktop_shell_order()
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Диспетчер задач'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def check_menu_item_exists(self, item_type):
        """Проверка существования пункта меню"""
        try:
            if item_type == "admin_menu":
                try:
                    clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Администрирование"
                except FileNotFoundError:
                    try:
                        clsid_path_hklm = r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, clsid_path_hklm) as key:
                            value, _ = winreg.QueryValueEx(key, "MUIVerb")
                            return value == "Администрирование"
                    except FileNotFoundError:
                        return False
                except:
                    return False

            elif item_type == "system_menu":
                try:
                    clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Система"
                except FileNotFoundError:
                    try:
                        clsid_path_hklm = r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, clsid_path_hklm) as key:
                            value, _ = winreg.QueryValueEx(key, "MUIVerb")
                            return value == "Система"
                    except FileNotFoundError:
                        return False
                except:
                    return False

            elif item_type == "regedit":
                try:
                    key_path = r'DesktopBackground\Shell\Regedit'
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Редактор реестра"
                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "restart_explorer":
                try:
                    key_path = r'DesktopBackground\Shell\RestartExplorer'
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Перезапустить explorer"
                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "power_menu":
                try:
                    power_keys = [
                        r'DesktopBackground\Shell\Hibernation',
                        r'DesktopBackground\Shell\Reboot',
                        r'DesktopBackground\Shell\Shutdown'
                    ]
                    for key_path in power_keys:
                        try:
                            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                                if "Hibernation" in key_path:
                                    return True
                                else:
                                    value, _ = winreg.QueryValueEx(key, "MUIVerb")
                                    return True
                        except:
                            continue
                    return False
                except:
                    return False
            elif item_type == "task_manager":
                try:
                    key_path = r'DesktopBackground\Shell\TaskManager'
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Диспетчер задач"
                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "copy_as_path":
                try:
                    key_path = r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                        value, reg_type = winreg.QueryValueEx(key, "")
                        return value == "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}"
                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "pin_to_start":
                try:
                    key_path = r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                        value, reg_type = winreg.QueryValueEx(key, "")
                        return value == "{470C0EBD-5D73-4d58-9CED-E91E22E23282}"
                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "pin_to_taskbar":
                try:
                    key_path = r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                        value, reg_type = winreg.QueryValueEx(key, "")
                        return value == "Taskband Pin"
                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "send_to_device":
                try:
                    blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path) as key:
                        try:
                            winreg.QueryValueEx(key, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}")
                            return False
                        except FileNotFoundError:
                            return True
                except FileNotFoundError:
                    return True
                except:
                    return False

            elif item_type == "paint3d":
                graphic_formats = ['bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                                   'obj', 'ply', 'png', 'stl', 'tif', 'tiff']

                for format_ext in graphic_formats[:3]:
                    try:
                        paint3d_key = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{format_ext}\\Shell\\3D Edit"
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key):
                            return True
                    except:
                        continue
                return False

            elif item_type == "compatibility":
                try:
                    blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path) as key:
                        try:
                            winreg.QueryValueEx(key, "{1d27f844-3a1f-4410-85ac-14651078412d}")
                            return False
                        except FileNotFoundError:
                            return True
                except FileNotFoundError:
                    return True
                except:
                    return True

            elif item_type == "defender":
                try:
                    defender_clsid_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
                    defender_clsid_no_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no'

                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path):
                            return True
                    except FileNotFoundError:
                        try:
                            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_no_path):
                                return False
                        except FileNotFoundError:
                            return False
                    except:
                        return False

                except:
                    return False

            elif item_type == "share_access":
                try:
                    blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path) as key:
                        try:
                            winreg.QueryValueEx(key, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}")
                            return False
                        except FileNotFoundError:
                            return True
                except:
                    return True

            elif item_type == "modern_sharing":
                try:
                    modern_sharing_paths = [
                        r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
                        r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
                        r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
                        r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
                        r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing'
                    ]

                    for path in modern_sharing_paths:
                        try:
                            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                                value, _ = winreg.QueryValueEx(key, "")
                                if value == "{e2bf9676-5f8f-435c-97eb-11607a5bedf7}":
                                    return True
                        except:
                            continue

                    return False

                except FileNotFoundError:
                    return False
                except:
                    return False

            elif item_type == "delete_folder_content":
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                        r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'):
                        return True
                except FileNotFoundError:
                    return False

            elif item_type == "copy_move":
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                        r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To'):
                        return True
                except FileNotFoundError:
                    return False

            elif item_type == "take_ownership":
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\*\shell\runas'):
                        return True
                except FileNotFoundError:
                    return False

            elif item_type == "previous_version":
                try:
                    blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path) as key:
                        try:
                            winreg.QueryValueEx(key, "{596AB062-B4D2-4215-9F74-E9109B0A8153}")
                            return False
                        except FileNotFoundError:
                            return True
                except FileNotFoundError:
                    return True
                except:
                    return True

        except:
            return False

        return False

    def add_copy_as_path(self):
        """Добавление пункта 'Копировать как путь'"""
        try:
            key_path = r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu'

            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}")

            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                    try:
                        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}")
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Копировать как путь'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_copy_as_path(self):
        """Удаление пункта 'Копировать как путь'"""
        try:
            key_path = r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu'

            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)

            self.update_menu_items_status()

        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError as pe:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE, key_path)
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Копировать как путь'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_pin_to_start(self):
        """Добавление пункта 'Закрепить на начальном экране'"""
        try:
            key_path = r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen'

            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "{470C0EBD-5D73-4d58-9CED-E91E22E23282}")

            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                    try:
                        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "{470C0EBD-5D73-4d58-9CED-E91E22E23282}")
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Закрепить на начальном экране'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_pin_to_start(self):
        """Удаление пункта 'Закрепить на начальном экране'"""
        try:
            key_path = r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen'

            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)

            self.update_menu_items_status()

        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError as pe:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE, key_path)
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Закрепить на начальном экране'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_pin_to_taskbar(self):
        """Добавление пункта 'Закрепить на Панели задач'"""
        try:
            key_path = r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}'

            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Taskband Pin")

            self.update_menu_items_status()

        except PermissionError as pe:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                    try:
                        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHine, key_path, 0, winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Taskband Pin")
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Закрепить на Панели задач'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_pin_to_taskbar(self):
        """Удаление пункта 'Закрепить на Панели задач'"""
        try:
            key_path = r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}'

            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)

            self.update_menu_items_status()

        except FileNotFoundError:
            self.update_menu_items_status()
        except PermissionError as pe:
            if self.take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if self.set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                    try:
                        self.delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE, key_path)
                        self.update_menu_items_status()
                    except:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Закрепить на Панели задач'")
                else:
                    messagebox.showerror("Ошибка", "Не удалось установить права доступа")
            else:
                messagebox.showerror("Ошибка", "Не удалось получить владение ключом реестра")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def delete_registry_key_recursive(self, hive, key_path):
        """Рекурсивное удаление ключа реестра со всеми подключами"""
        try:
            with winreg.OpenKey(hive, key_path) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey_path = key_path + "\\" + subkey_name
                        self.delete_registry_key_recursive(hive, subkey_path)
                    except OSError:
                        break
                    i += 1
        except FileNotFoundError:
            return
        except:
            pass

        try:
            winreg.DeleteKey(hive, key_path)
        except:
            pass

    def add_send_to_device(self):
        """Добавление пункта 'Передать на устройство'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'

            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}")
                except FileNotFoundError:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_send_to_device(self):
        """Удаление пункта 'Передать на устройство'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'

            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}", 0,
                                  winreg.REG_SZ, "Передать на устройство")

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_paint3d_edit(self):
        """Добавление пункта 'Изменить с помощью Paint 3D' для всех графических форматов"""
        try:
            graphic_formats = ['bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                               'obj', 'ply', 'png', 'stl', 'tif', 'tiff']

            for format_ext in graphic_formats:
                try:
                    shell_key_path = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{format_ext}\\Shell"

                    try:
                        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, shell_key_path)
                    except:
                        pass

                    paint3d_key = f"{shell_key_path}\\3D Edit"
                    try:
                        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key)
                    except:
                        pass

                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key, 0,
                                            winreg.KEY_SET_VALUE) as key:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                                              "@%SystemRoot%\\system32\\mspaint.exe,-59500")
                    except:
                        pass

                    command_key = f"{paint3d_key}\\command"
                    try:
                        winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key)
                    except:
                        pass

                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_key, 0,
                                            winreg.KEY_SET_VALUE) as cmd_key:
                            command_value = '%SystemRoot%\\system32\\mspaint.exe "%1" /ForceBootstrapPaint3D'
                            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_EXPAND_SZ,
                                              command_value)
                    except:
                        pass

                except PermissionError as pe:
                    try:
                        full_path = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{format_ext}"
                        if self.take_ownership_registry_key_simple(full_path):
                            if self.set_registry_key_permissions_simple(full_path):
                                try:
                                    paint3d_key = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{format_ext}\\Shell\\3D Edit"
                                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key)

                                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key, 0,
                                                        winreg.KEY_SET_VALUE) as key:
                                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                                                          "@%SystemRoot%\\system32\\mspaint.exe,-59500")

                                    command_key = f"{paint3d_key}\\command"
                                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key)

                                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_key, 0,
                                                        winreg.KEY_SET_VALUE) as cmd_key:
                                        command_value = '%SystemRoot%\\system32\\mspaint.exe "%1" /ForceBootstrapPaint3D'
                                        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_EXPAND_SZ,
                                                          command_value)
                                except:
                                    pass
                    except:
                        pass
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_paint3d_edit(self):
        """Удаление пункта 'Изменить с помощью Paint 3D' для всех графических форматов"""
        try:
            graphic_formats = ['bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                               'obj', 'ply', 'png', 'stl', 'tif', 'tiff']

            for format_ext in graphic_formats:
                try:
                    paint3d_key = f"SOFTWARE\\Classes\\SystemFileAssociations\\.{format_ext}\\Shell\\3D Edit"

                    key_exists = False
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key):
                            key_exists = True
                    except FileNotFoundError:
                        key_exists = False

                    if key_exists:
                        command_key = f"{paint3d_key}\\command"
                        try:
                            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, command_key)
                        except FileNotFoundError:
                            pass
                        except PermissionError as pe:
                            if self.take_ownership_registry_key_simple(command_key):
                                if self.set_registry_key_permissions_simple(command_key):
                                    try:
                                        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHine, command_key)
                                    except:
                                        pass

                        try:
                            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key)
                        except PermissionError as pe:
                            try:
                                if self.take_ownership_registry_key_simple(paint3d_key):
                                    if self.set_registry_key_permissions_simple(paint3d_key):
                                        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, paint3d_key)
                            except:
                                pass
                        except:
                            pass

                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_defender_scan(self):
        """Добавление проверки Windows Defender"""
        try:
            defender_clsid_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
            defender_clsid_no_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no'

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_no_path):
                    if self.rename_registry_key(winreg.HKEY_LOCAL_MACHINE, defender_clsid_no_path, defender_clsid_path):
                        self.update_menu_items_status()
                        return
            except:
                pass

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path):
                    self.update_menu_items_status()
                    return
            except:
                pass

            try:
                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path)
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Microsoft.Windows.Defender.IconHandler")

                inproc_path = defender_clsid_path + r'\InprocServer32'
                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, inproc_path)
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, inproc_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_EXPAND_SZ,
                                      "%SystemRoot%\\System32\\WindowsDefender\\ShellExt.dll")
                    winreg.SetValueEx(key, "ThreadingModel", 0, winreg.REG_SZ, "Apartment")

                self.update_menu_items_status()

            except:
                messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта")

        except:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта")

    def remove_defender_scan(self):
        """Удаление проверки Windows Defender"""
        try:
            defender_clsid_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
            defender_clsid_no_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no'

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path):
                    if self.rename_registry_key(winreg.HKEY_LOCAL_MACHINE, defender_clsid_path, defender_clsid_no_path):
                        self.update_menu_items_status()
                        return True
                    else:
                        messagebox.showerror("Ошибка", "Не удалось удалить пункт")
                        return False
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_clsid_no_path):
                        self.update_menu_items_status()
                        return True
                except FileNotFoundError:
                    self.update_menu_items_status()
                    return True
            except:
                messagebox.showerror("Ошибка", "Ошибка при удалении пункта")
                return False

        except:
            messagebox.showerror("Ошибка", "Ошибка при удалении пункта")
            return False

    def add_modern_sharing(self):
        """Добавление современного меню отправки (Metro)"""
        try:
            modern_sharing_paths = [
                (r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
                 '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'),
                (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
                 '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'),
                (r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
                 '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'),
                (r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
                 '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'),
                (r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing',
                 '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}')
            ]

            for key_path, clsid_value in modern_sharing_paths:
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)

                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHine, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, clsid_value)

                except PermissionError as pe:
                    if self.take_ownership_registry_key_simple(key_path):
                        if self.set_registry_key_permissions_simple(key_path):
                            try:
                                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0,
                                                    winreg.KEY_SET_VALUE) as key:
                                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, clsid_value)
                            except:
                                pass
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_modern_sharing(self):
        """Удаление современного меню отправки (Metro)"""
        try:
            modern_sharing_paths = [
                r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
                r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing'
            ]

            for key_path in modern_sharing_paths:
                try:
                    try:
                        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    except FileNotFoundError:
                        pass
                    except PermissionError as pe:
                        if self.take_ownership_registry_key_simple(key_path):
                            if self.set_registry_key_permissions_simple(key_path):
                                try:
                                    self.delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE, key_path)
                                except:
                                    pass
                    except:
                        pass

                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_delete_folder_content(self):
        """Добавление удаления содержимого папки (HKLM)"""
        try:
            delete_folder_content_keys = [
                r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent',
            ]

            for key_path in delete_folder_content_keys:
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Удалить содержимое папки")
                        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "shell32.dll,-240")

                    command_key = key_path + r'\command'
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_key, 0, winreg.KEY_SET_VALUE) as cmd_key:
                        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ,
                                          'cmd.exe /c "cd /d \"%1\" && del /q /f /s *.*"')
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_delete_folder_content(self):
        """Удаление удаления содержимого папки (HKLM)"""
        try:
            delete_folder_content_keys = [
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent\command'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'),
            ]

            for hive, key_path in delete_folder_content_keys:
                try:
                    winreg.DeleteKey(hive, key_path)
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_copy_move_menu(self):
        """Добавление копирования/перемещения в папку (HKLM)"""
        try:
            copy_move_keys = [
                r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
                r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To',
            ]

            for key_path in copy_move_keys:
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        if "Copy To" in key_path:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "{C2FBB630-2971-11D1-A18C-00C04FD75D13}")
                        else:
                            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "{C2FBB631-2971-11D1-A18C-00C04FD75D13}")
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении меню: {str(e)}")

    def remove_copy_move_menu(self):
        """Удаление копирования/перемещения в папку (HKLM)"""
        try:
            copy_move_keys = [
                (winreg.HKEY_LOCAL_MACHINE,
                 r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To'),
                (winreg.HKEY_LOCAL_MACHINE,
                 r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To'),
            ]

            for hive, key_path in copy_move_keys:
                try:
                    winreg.DeleteKey(hive, key_path)
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении меню: {str(e)}")

    def add_take_ownership(self):
        """Добавление 'Стать владельцем' (HKLM)"""
        try:
            take_ownership_keys = [
                r'SOFTWARE\Classes\*\shell\runas',
                r'SOFTWARE\Classes\Directory\shell\runas',
            ]

            for key_path in take_ownership_keys:
                try:
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Стать владельцем")
                        winreg.SetValueEx(key, "HasLUAShield", 0, winreg.REG_SZ, "")

                    command_key = key_path + r'\command'
                    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_key)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_key, 0, winreg.KEY_SET_VALUE) as cmd_key:
                        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ,
                                          'cmd.exe /c "takeown /f \"%1\" && icacls \"%1\" /grant administrators:F"')
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_take_ownership(self):
        """Удаление 'Стать владельцем' (HKLM)"""
        try:
            take_ownership_keys = [
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\*\shell\runas\command'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\*\shell\runas'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\Directory\shell\runas\command'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\Directory\shell\runas'),
            ]

            for hive, key_path in take_ownership_keys:
                try:
                    winreg.DeleteKey(hive, key_path)
                except:
                    pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_compatibility(self):
        """Добавление пункта 'Исправление проблем с совместимостью'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                    try:
                        winreg.DeleteValue(key, "{1d27f844-3a1f-4410-85ac-14651078412d}")
                    except:
                        pass
            except:
                pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_compatibility(self):
        """Удаление 'Исправление проблем с совместимостью'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            try:
                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "{1d27f844-3a1f-4410-85ac-14651078412d}", 0, winreg.REG_SZ,
                                      "Исправление проблем с совместимостью")
            except:
                pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_previous_version(self):
        """Добавление пункта 'Восстановить прежнюю версию'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                    try:
                        winreg.DeleteValue(key, "{596AB062-B4D2-4215-9F74-E9109B0A8153}")
                    except:
                        pass
            except:
                pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {str(e)}")

    def remove_previous_version(self):
        """Удаление 'Восстановить прежнюю версию'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            try:
                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "{596AB062-B4D2-4215-9F74-E9109B0A8153}", 0, winreg.REG_SZ,
                                      "Восстановить прежнюю версию")
            except:
                pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {str(e)}")

    def add_share_access(self):
        """Добавление меню 'Предоставить доступ к'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                    try:
                        winreg.DeleteValue(key, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}")
                    except:
                        pass
            except:
                pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при добавлении меню: {str(e)}")

    def remove_share_access(self):
        """Удаление меню 'Предоставить доступ к'"""
        try:
            blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
            try:
                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, blocked_path)
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}", 0, winreg.REG_SZ, "")
            except:
                pass

            self.update_menu_items_status()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при удалении меню: {str(e)}")

    def create_menu_item_button(self, parent, item):
        """Создание кнопки с индикатором состояния"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack()

        indicator_canvas = tk.Canvas(btn_frame, width=20, height=20, highlightthickness=0)
        indicator_canvas.pack(side=tk.LEFT, padx=(0, 5))

        indicator_canvas.create_oval(2, 2, 18, 18, outline='#e74c3c', fill='#e74c3c')
        indicator_canvas.create_text(10, 10, text="✗", fill="white", font=('Arial', 10, 'bold'))

        item["indicator_canvas"] = indicator_canvas
        item["indicator_circle"] = indicator_canvas.find_all()[0]
        item["indicator_text"] = indicator_canvas.find_all()[1]

        btn = ttk.Button(
            btn_frame,
            text=item["name"],
            width=50,
            command=lambda i=item: self.toggle_menu_item(i)
        )
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=2)

        return btn

    def update_indicator(self, item, is_enabled):
        """Обновление индикатора состояния"""
        if 'indicator_canvas' in item and item['indicator_canvas']:
            canvas = item['indicator_canvas']
            if is_enabled:
                canvas.itemconfig(item['indicator_circle'], outline='#27ae60', fill='#27ae60')
                canvas.itemconfig(item['indicator_text'], text="✓")
            else:
                canvas.itemconfig(item['indicator_circle'], outline='#e74c3c', fill='#e74c3c')
                canvas.itemconfig(item['indicator_text'], text="✗")

    def create_tooltip(self, widget, text):
        """Создание подсказки для виджета"""

        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            label = ttk.Label(tooltip, text=text, background="#ffffe0", relief='solid', borderwidth=1)
            label.pack()
            widget.tooltip = tooltip

        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def restart_explorer(self):
        """Перезапуск проводника Windows"""
        try:
            subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True)
            subprocess.Popen("explorer.exe")
        except:
            messagebox.showerror("Ошибка", "Ошибка при перезагрузке проводника")

    def set_default_settings(self):
        """Установка настроек по умолчанию"""
        try:
            for item in self.menu_item_buttons:
                item_data = item["data"]

                current_state = item_data["check_command"]()

                if item_data.get("default_enabled", True) and not current_state:
                    try:
                        item_data["add_command"]()
                    except:
                        pass
                elif not item_data.get("default_enabled", True) and current_state:
                    try:
                        item_data["remove_command"]()
                    except:
                        pass

            self.update_power_menu_icons()
            self.update_menu_items_status()
            self.update_desktop_shell_order()

            messagebox.showinfo("Успех", "Настройки сброшены по умолчанию!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сбросе настроек: {str(e)}")

    def toggle_menu_item(self, item):
        """Переключение состояния пункта меню"""
        try:
            exists = item["check_command"]()

            if exists:
                item["remove_command"]()
                item["enabled"] = False
            else:
                item["add_command"]()
                item["enabled"] = True

            self.update_indicator(item, not exists)
            self.update_menu_items_status()

            if item["name"] in ["Выключение/Перезагрузка/Гибернация", "Диспетчер задач", "Перезапустить explorer",
                                "Редактор реестра"]:
                self.update_desktop_shell_order()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта: {str(e)}")

    def create_new_menu_tab(self, parent):
        """Создание вкладки меню Создать"""

        main_container = ttk.Frame(parent)
        main_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        title_label = ttk.Label(
            scrollable_frame,
            text="Управление пунктами меню 'Создать'",
            font=('Arial', 12, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        control_frame = ttk.Frame(scrollable_frame)
        control_frame.pack(pady=5)

        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)

        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        self.new_menu_toggle_all_btn = ttk.Button(
            buttons_frame,
            text="",
            command=self.toggle_all_new_menu_items,
            width=15
        )
        self.new_menu_toggle_all_btn.pack(side=tk.LEFT, padx=5, pady=5)

        separator = ttk.Separator(scrollable_frame, orient='horizontal')
        separator.pack(fill=tk.X, padx=10, pady=10)

        self.new_menu_items_frame = ttk.Frame(scrollable_frame)
        self.new_menu_items_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.refresh_new_menu_list()

    def toggle_all_new_menu_items(self):
        """Переключение всех пунктов меню Создать"""
        try:
            items = self.get_new_menu_items()
            if not items:
                return

            # Проверяем текущее состояние всех пунктов
            all_enabled = all(item["enabled"] for item in items)
            all_disabled = all(not item["enabled"] for item in items)

            # Определяем действие: если есть хоть один включенный - отключаем всё, иначе включаем всё
            enable_all = all_disabled

            for item in items:
                try:
                    if enable_all:
                        # Включаем все отключенные пункты
                        if not item["enabled"]:
                            if item["ext"] == '.folder':
                                # Особый случай для папки (работа с HKCR)
                                self.enable_folder_menu_item(item.get("path"))
                            else:
                                # Для остальных пунктов (работа с HKLM)
                                self.enable_new_menu_item(item["ext"], item.get("path"))
                    else:
                        # Отключаем все включенные пункты
                        if item["enabled"]:
                            if item["ext"] == '.folder':
                                # Особый случай для папки
                                self.disable_folder_menu_item(item.get("path"))
                            else:
                                # Для остальных пунктов
                                self.disable_new_menu_item(item["ext"], item.get("path"))
                except Exception as e:
                    print(f"Ошибка при переключении пункта {item['name']}: {e}")
                    # Показываем ошибку только для конкретного пункта
                    messagebox.showerror("Ошибка", f"Не удалось переключить пункт '{item['name']}': {str(e)}")

            # Обновляем список
            self.refresh_new_menu_list()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пунктов: {str(e)}")

    def refresh_new_menu_list(self):
        """Обновление списка пунктов меню Создать"""
        try:
            for widget in self.new_menu_items_frame.winfo_children():
                widget.destroy()

            items = self.get_new_menu_items()

            if not items:
                no_items_label = ttk.Label(
                    self.new_menu_items_frame,
                    text="Пункты меню 'Создать' не найдены",
                    font=('Arial', 10),
                    foreground='#7f8c8d'
                )
                no_items_label.pack(pady=20)
                return

            all_enabled = all(item["enabled"] for item in items)
            all_disabled = all(not item["enabled"] for item in items)

            if all_disabled:
                self.new_menu_toggle_all_btn.config(text="Включить всё")
            else:
                self.new_menu_toggle_all_btn.config(text="Отключить всё")

            header_frame = ttk.Frame(self.new_menu_items_frame)
            header_frame.pack(fill=tk.X, pady=(0, 5))

            ttk.Label(header_frame, text="Название пункта", font=('Arial', 9, 'bold'), width=30).pack(side=tk.LEFT,
                                                                                                      padx=5)
            ttk.Label(header_frame, text="Статус", font=('Arial', 9, 'bold'), width=15).pack(side=tk.LEFT, padx=5)
            ttk.Label(header_frame, text="Действие", font=('Arial', 9, 'bold'), width=15).pack(side=tk.LEFT, padx=0)

            self.new_menu_widgets = []
            for item in items:
                self.create_new_menu_item_row(item)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении списка: {str(e)}")

    def create_new_menu_item_row(self, item):
        """Создание строки с пунктом меню Создать"""
        try:
            row_frame = ttk.Frame(self.new_menu_items_frame, relief='groove', borderwidth=1)
            row_frame.pack(fill=tk.X, pady=2, padx=5)

            name_label = ttk.Label(
                row_frame,
                text=item["name"],
                width=29,
                anchor='w'
            )
            name_label.pack(side=tk.LEFT, padx=5, pady=5)

            status_text = "✅ Включен" if item["enabled"] else "❌ Отключен"
            status_label = ttk.Label(
                row_frame,
                text=status_text,
                width=15,
                foreground='#27ae60' if item["enabled"] else '#e74c3c'
            )
            status_label.pack(side=tk.LEFT, padx=5, pady=5)

            action_text = "Отключить" if item["enabled"] else "Включить"
            action_btn = ttk.Button(
                row_frame,
                text=action_text,
                width=15,
                command=lambda ext=item["ext"], path=item.get("path", f'SOFTWARE\\Classes\\{item["ext"]}\\ShellNew'):
                self.toggle_new_menu_item(ext, path)
            )
            action_btn.pack(side=tk.LEFT, padx=5, pady=5)

            self.new_menu_widgets.append({
                "frame": row_frame,
                "name": name_label,
                "status": status_label,
                "button": action_btn,
                "ext": item["ext"]
            })

        except:
            pass

    def get_new_menu_items(self):
        """Получение списка пунктов меню Создать (используем HKLM)"""
        items = []
        try:
            # Добавляем пункт "Папка" в начало списка
            folder_ext = '.folder'  # Виртуальное расширение для папки
            folder_path = r'Folder\ShellNew'

            # Проверяем состояние папки (работа с HKCR)
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path):
                    items.append({
                        "ext": folder_ext,
                        "name": "Папка",
                        "enabled": True,
                        "path": folder_path
                    })
            except FileNotFoundError:
                folder_no_path = folder_path + '_no'
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_no_path):
                        items.append({
                            "ext": folder_ext,
                            "name": "Папка",
                            "enabled": False,
                            "path": folder_path
                        })
                except FileNotFoundError:
                    # Проверяем, может быть вообще нет записи для папки
                    try:
                        # Проверяем, есть ли хотя бы ключ Folder
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'Folder'):
                            items.append({
                                "ext": folder_ext,
                                "name": "Папка",
                                "enabled": False,
                                "path": folder_path
                            })
                    except FileNotFoundError:
                        # Ключа Folder нет вообще
                        pass

            target_extensions = {
                '.bmp': r'SOFTWARE\Classes\.bmp\ShellNew',
                '.contact': r'SOFTWARE\Classes\.contact\ShellNew',
                '.lnk': r'SOFTWARE\Classes\.lnk\ShellNew',
                '.rtf': r'SOFTWARE\Classes\.rtf\ShellNew',
                '.txt': r'SOFTWARE\Classes\.txt\ShellNew',
                '.zip': r'SOFTWARE\Classes\.zip\CompressedFolder\ShellNew'
            }

            office_extensions = self.get_office_extensions()

            for ext, path in office_extensions.items():
                target_extensions[ext] = path

            for ext, base_path in target_extensions.items():
                if not base_path:
                    continue

                shellnew_path = base_path
                shellnew_no_path = base_path + '_no'

                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path):
                        items.append({
                            "ext": ext,
                            "name": self.get_file_type_name(ext),
                            "enabled": True,
                            "path": base_path
                        })
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_no_path):
                            items.append({
                                "ext": ext,
                                "name": self.get_file_type_name(ext),
                                "enabled": False,
                                "path": base_path
                            })
                    except FileNotFoundError:
                        if ext in office_extensions:
                            try:
                                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Classes\\{ext}"):
                                    items.append({
                                        "ext": ext,
                                        "name": self.get_file_type_name(ext),
                                        "enabled": False,
                                        "path": base_path
                                    })
                            except FileNotFoundError:
                                if ext == '.docx':
                                    try:
                                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                                            f"SOFTWARE\\Classes\\{ext}\\Word.Document.12"):
                                            items.append({
                                                "ext": ext,
                                                "name": self.get_file_type_name(ext),
                                                "enabled": False,
                                                "path": base_path
                                            })
                                    except FileNotFoundError:
                                        pass
                                elif ext == '.xlsx':
                                    try:
                                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                                            f"SOFTWARE\\Classes\\{ext}\\Excel.Sheet.12"):
                                            items.append({
                                                "ext": ext,
                                                "name": self.get_file_type_name(ext),
                                                "enabled": False,
                                                "path": base_path
                                            })
                                    except FileNotFoundError:
                                        pass
                                elif ext == '.pptx':
                                    try:
                                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                                            f"SOFTWARE\\Classes\\{ext}\\PowerPoint.Show.12"):
                                            items.append({
                                                "ext": ext,
                                                "name": self.get_file_type_name(ext),
                                                "enabled": False,
                                                "path": base_path
                                            })
                                    except FileNotFoundError:
                                        pass
                except:
                    pass

        except Exception as e:
            print(f"Ошибка в get_new_menu_items: {e}")

        return items

    def get_office_extensions(self):
        """Получение списка расширений Microsoft Office, если он установлен"""
        office_extensions = {}

        office_installed = False

        try:
            office_versions = [
                r'SOFTWARE\Microsoft\Office\16.0',
                r'SOFTWARE\Microsoft\Office\15.0',
                r'SOFTWARE\Microsoft\Office\14.0',
                r'SOFTWARE\Microsoft\Office\12.0',
                r'SOFTWARE\Microsoft\Office\11.0'
            ]

            for version_path in office_versions:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, version_path):
                        office_installed = True
                        break
                except:
                    continue

            if not office_installed:
                for version_path in office_versions:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, version_path):
                            office_installed = True
                            break
                    except:
                        continue

            if not office_installed:
                try:
                    uninstall_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key) as key:
                        i = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(key, i)
                                subkey_path = f"{uninstall_key}\\{subkey_name}"
                                try:
                                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey_path) as subkey:
                                        display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                                        display_name_str = str(display_name)
                                        if "Microsoft Office" in display_name_str or "Microsoft Word" in display_name_str or "Microsoft Excel" in display_name_str or "Microsoft PowerPoint" in display_name_str:
                                            office_installed = True
                                            break
                                except:
                                    pass
                                i += 1
                            except OSError:
                                break
                except:
                    pass

        except:
            pass

        if office_installed:
            office_exts = ['.docx', '.xlsx', '.pptx', '.pub']

            for ext in office_exts:
                possible_paths = [
                    f'SOFTWARE\\Classes\\{ext}\\ShellNew',
                    f'SOFTWARE\\Classes\\SystemFileAssociations\\{ext}\\ShellNew'
                ]

                if ext == '.docx':
                    possible_paths.insert(0, r'SOFTWARE\Classes\.docx\Word.Document.12\ShellNew')
                elif ext == '.xlsx':
                    possible_paths.insert(0, r'SOFTWARE\Classes\.xlsx\Excel.Sheet.12\ShellNew')
                elif ext == '.pptx':
                    possible_paths.insert(0, r'SOFTWARE\Classes\.pptx\PowerPoint.Show.12\ShellNew')
                elif ext == '.pub':
                    possible_paths.insert(0, r'SOFTWARE\Classes\.pub\Publisher.Document.12\ShellNew')

                found_path = None
                for path in possible_paths:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                            found_path = path
                            break
                    except:
                        continue

                if found_path:
                    office_extensions[ext] = found_path
                else:
                    office_extensions[ext] = f'SOFTWARE\\Classes\\{ext}\\ShellNew'

        return office_extensions

    def get_file_type_name(self, ext):
        """Получение читаемого имени для расширения файла"""
        basic_names = {
            '.folder': 'Папка',
            '.bmp': 'Рисунок BMP',
            '.contact': 'Контакт',
            '.lnk': 'Ярлык',
            '.rtf': 'Документ RTF',
            '.txt': 'Текстовый документ',
            '.zip': 'Архив ZIP',
            '.docx': 'Документ Word',
            '.xlsx': 'Книга Excel',
            '.pptx': 'Презентация PowerPoint',
            '.pub': 'Публикация Publisher'
        }

        if ext in basic_names:
            return basic_names[ext]
        else:
            clean_ext = ext.lstrip('.')
            return f"Файл {clean_ext.upper()}"

    def toggle_new_menu_item(self, ext, base_path=None):
        """Переключение состояния пункта меню Создать (HKLM)"""
        try:
            # Особый случай для папки
            if ext == '.folder':
                if base_path is None:
                    base_path = r'Folder\ShellNew'

                return self.toggle_folder_menu_item(base_path)

            if base_path is None:
                base_path = f'SOFTWARE\\Classes\\{ext}\\ShellNew'

            shellnew_path = base_path
            shellnew_no_path = base_path + '_no'

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path):
                    if not self.disable_new_menu_item(ext, base_path):
                        messagebox.showerror("Ошибка", f"Не удалось отключить пункт '{self.get_file_type_name(ext)}'")
                    self.refresh_new_menu_list()
                    return
            except FileNotFoundError:
                pass

            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_no_path):
                    if not self.enable_new_menu_item(ext, base_path):
                        messagebox.showerror("Ошибка", f"Не удалось включить пункт '{self.get_file_type_name(ext)}'")
                    self.refresh_new_menu_list()
                    return
            except FileNotFoundError:
                if ext in ['.docx', '.xlsx', '.pptx', '.pub']:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Classes\\{ext}"):
                            pass
                        if not self.enable_new_menu_item(ext, base_path):
                            messagebox.showerror("Ошибка", f"Не удалось создать пункт '{self.get_file_type_name(ext)}'")
                        self.refresh_new_menu_list()
                        return
                    except FileNotFoundError:
                        if ext == '.docx':
                            alt_path = r'SOFTWARE\Classes\.docx\Word.Document.12'
                        elif ext == '.xlsx':
                            alt_path = r'SOFTWARE\Classes\.xlsx\Excel.Sheet.12'
                        elif ext == '.pptx':
                            alt_path = r'SOFTWARE\Classes\.pptx\PowerPoint.Show.12'
                        elif ext == '.pub':
                            alt_path = r'SOFTWARE\Classes\.pub\Publisher.Document.12'
                        else:
                            alt_path = None

                        if alt_path:
                            try:
                                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, alt_path):
                                    if not self.enable_new_menu_item(ext, base_path):
                                        messagebox.showerror("Ошибка",
                                                             f"Не удалось создать пункт '{self.get_file_type_name(ext)}'")
                                    self.refresh_new_menu_list()
                                    return
                            except FileNotFoundError:
                                messagebox.showinfo("Информация",
                                                    f"Microsoft Office не установлен или тип файла {ext} не ассоциирован")
                                return
                        else:
                            messagebox.showinfo("Информация",
                                                f"Microsoft Office не установлен или тип файла {ext} не ассоциирован")
                            return
                else:
                    if not self.enable_new_menu_item(ext, base_path):
                        messagebox.showerror("Ошибка", f"Не удалось создать пункт '{self.get_file_type_name(ext)}'")
                    self.refresh_new_menu_list()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта: {str(e)}")

    def toggle_folder_menu_item(self, base_path):
        """Переключение пункта 'Папка' в меню Создать"""
        try:
            shellnew_path = base_path
            shellnew_no_path = base_path + '_no'

            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_path):
                    if not self.disable_folder_menu_item(base_path):
                        messagebox.showerror("Ошибка", "Не удалось отключить пункт 'Папка'")
                    self.refresh_new_menu_list()
                    return True
            except FileNotFoundError:
                pass

            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_no_path):
                    if not self.enable_folder_menu_item(base_path):
                        messagebox.showerror("Ошибка", "Не удалось включить пункт 'Папка'")
                    self.refresh_new_menu_list()
                    return True
            except FileNotFoundError:
                if not self.enable_folder_menu_item(base_path):
                    messagebox.showerror("Ошибка", "Не удалось создать пункт 'Папка'")
                self.refresh_new_menu_list()
                return True

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта 'Папка': {str(e)}")
            return False

    def enable_folder_menu_item(self, base_path):
        """Включение пункта 'Папка' в меню Создать"""
        try:
            if base_path is None:
                base_path = r'Folder\ShellNew'

            shellnew_path = base_path
            shellnew_no_path = base_path + '_no'

            # Сначала пытаемся переименовать из _no в обычный
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_no_path):
                    if self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, shellnew_no_path, shellnew_path):
                        return True
                    else:
                        raise Exception("Не удалось переименовать ключ папки")
            except FileNotFoundError:
                pass

            # Если нет _no, проверяем может ключ уже существует
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_path):
                    # Ключ уже существует и включен
                    return True
            except FileNotFoundError:
                # Если ключа нет вообще, создаем минимальную запись для папки
                try:
                    winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, shellnew_path)
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                    return True
                except Exception as e:
                    raise Exception(f"Не удалось создать ключ папки: {e}")

        except Exception as e:
            print(f"Ошибка в enable_folder_menu_item: {e}")
            return False

    def disable_folder_menu_item(self, base_path):
        """Отключение пункта 'Папка' в меню Создать"""
        try:
            if base_path is None:
                base_path = r'Folder\ShellNew'

            shellnew_path = base_path
            shellnew_no_path = base_path + '_no'

            # Проверяем, существует ли уже отключенная версия
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_no_path):
                    # Уже отключено
                    return True
            except FileNotFoundError:
                pass

            # Пытаемся переименовать обычный ключ в _no
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shellnew_path):
                    if self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, shellnew_path, shellnew_no_path):
                        return True
                    else:
                        raise Exception("Не удалось переименовать ключ папки")
            except FileNotFoundError:
                # Если ключа нет, значит он уже отключен или не существует
                return True

        except Exception as e:
            print(f"Ошибка в disable_folder_menu_item: {e}")
            return False

    def enable_new_menu_item(self, ext, base_path=None):
        """Включение пункта в меню Создать (HKLM)"""
        try:
            if base_path is None:
                base_path = f'SOFTWARE\\Classes\\{ext}\\ShellNew'

            shellnew_path = base_path
            shellnew_no_path = base_path + '_no'

            try:
                # Пытаемся переименовать из _no в обычный
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_no_path):
                    if self.rename_registry_key(winreg.HKEY_LOCAL_MACHINE, shellnew_no_path, shellnew_path):
                        return True
                    else:
                        raise Exception(f"Не удалось переименовать ключ для {ext}")
            except FileNotFoundError:
                # Если нет _no, проверяем может ключ уже существует
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path):
                        # Ключ уже существует и включен
                        return True
                except FileNotFoundError:
                    # Если ключа нет вообще, создаем его
                    try:
                        if ext in ['.docx', '.xlsx', '.pptx', '.pub']:
                            return self.create_office_shellnew_entry(ext, shellnew_path)
                        else:
                            return self.create_basic_shellnew_entry(ext, shellnew_path)
                    except Exception as e:
                        raise Exception(f"Не удалось создать ключ для {ext}: {e}")

        except Exception as e:
            print(f"Ошибка в enable_new_menu_item для {ext}: {e}")
            return False

    def disable_new_menu_item(self, ext, base_path=None):
        """Отключение пункта в меню Создать (HKLM)"""
        try:
            if base_path is None:
                base_path = f'SOFTWARE\\Classes\\{ext}\\ShellNew'

            shellnew_path = base_path
            shellnew_no_path = base_path + '_no'

            # Проверяем, существует ли уже отключенная версия
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_no_path):
                    # Уже отключено
                    return True
            except FileNotFoundError:
                pass

            # Пытаемся переименовать обычный ключ в _no
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path):
                    if self.rename_registry_key(winreg.HKEY_LOCAL_MACHINE, shellnew_path, shellnew_no_path):
                        return True
                    else:
                        raise Exception(f"Не удалось переименовать ключ для {ext}")
            except FileNotFoundError:
                # Если ключа нет, значит он уже отключен или не существует
                return True

        except Exception as e:
            print(f"Ошибка в disable_new_menu_item для {ext}: {e}")
            return False

    def create_office_shellnew_entry(self, ext, shellnew_path):
        """Создание записи ShellNew для файлов Office"""
        try:
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path)

            template_names = {
                '.docx': 'Word.Document.12',
                '.xlsx': 'Excel.Sheet.12',
                '.pptx': 'PowerPoint.Show.12',
                '.pub': 'Publisher.Document.12'
            }

            template_name = template_names.get(ext, '')

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")

                if ext in ['.docx', '.xlsx', '.pptx']:
                    winreg.SetValueEx(key, "FileName", 0, winreg.REG_SZ, template_name)

            return True

        except:
            try:
                command_path = shellnew_path + r'\command'
                winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, command_path)

                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, command_path, 0, winreg.KEY_SET_VALUE) as cmd_key:
                    cmd_value = f'cmd.exe /c echo. > "%1{ext}"'
                    winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd_value)

                return True

            except:
                return False

    def create_basic_shellnew_entry(self, ext, shellnew_path):
        """Создание базовой записи ShellNew"""
        try:
            winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path)

            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path, 0, winreg.KEY_SET_VALUE) as key:
                # Для разных типов файлов разные параметры
                if ext == '.bmp':
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                elif ext == '.txt':
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                elif ext == '.rtf':
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                elif ext == '.lnk':
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                elif ext == '.contact':
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                elif ext == '.zip':
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                else:
                    winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")

            return True
        except Exception as e:
            print(f"Ошибка создания базового ShellNew для {ext}: {e}")
            return False

    def create_sendto_menu_tab(self, parent):
        """Создание вкладки меню Отправить"""

        title_label = ttk.Label(
            parent,
            text="Управление пунктами меню 'Отправить'",
            font=('Arial', 12, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        control_frame = ttk.Frame(parent)
        control_frame.pack(pady=5)

        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)

        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        self.sendto_toggle_all_btn = ttk.Button(
            buttons_frame,
            text="",
            command=self.toggle_all_sendto_items,
            width=15
        )
        self.sendto_toggle_all_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.toggle_selected_btn = ttk.Button(
            buttons_frame,
            text="Включить",
            command=self.toggle_selected_sendto_item,
            width=15
        )
        self.toggle_selected_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.sendto_listbox = tk.Listbox(parent, height=15, font=('Arial', 9))
        self.sendto_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.sendto_listbox.bind('<<ListboxSelect>>', self.on_sendto_selection)

        self.refresh_sendto_list()

    def toggle_all_sendto_items(self):
        """Переключение всех пунктов меню Отправить"""
        try:
            sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'SendTo')

            if not os.path.exists(sendto_path):
                messagebox.showerror("Ошибка", "Папка SendTo не найдена")
                return

            items = os.listdir(sendto_path)

            has_enabled_items = False
            has_disabled_items = False

            for item in items:
                if item.lower() == "desktop.ini":
                    continue

                if item.endswith('.DeskLink_no'):
                    has_disabled_items = True
                else:
                    has_enabled_items = True

            all_disabled = has_disabled_items and not has_enabled_items

            for item in items:
                item_path = os.path.join(sendto_path, item)

                if item.lower() == "desktop.ini":
                    continue

                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if all_disabled:
                        if item.endswith('.DeskLink_no'):
                            original_name = item[:-12]
                            disabled_path = item_path
                            enabled_path = os.path.join(sendto_path, original_name)

                            try:
                                if os.path.exists(disabled_path) and not os.path.exists(enabled_path):
                                    os.rename(disabled_path, enabled_path)
                            except:
                                pass
                    else:
                        if not item.endswith('.DeskLink_no'):
                            disabled_path = os.path.join(sendto_path, f"{item}.DeskLink_no")
                            enabled_path = item_path

                            try:
                                if os.path.exists(enabled_path) and not os.path.exists(disabled_path):
                                    os.rename(enabled_path, disabled_path)
                            except:
                                pass

            self.refresh_sendto_list()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пунктов: {str(e)}")

    def on_sendto_selection(self, event):
        """Обработчик выбора пункта в меню Отправить"""
        selection = self.sendto_listbox.curselection()
        if selection:
            item_text = self.sendto_listbox.get(selection[0])
            if "❌" in item_text:
                self.toggle_selected_btn.config(text="Включить")
            else:
                self.toggle_selected_btn.config(text="Отключить")

    def toggle_selected_sendto_item(self):
        """Переключение состояния выделенного пункта в меню Отправить"""
        try:
            selection = self.sendto_listbox.curselection()
            if not selection:
                messagebox.showinfo("Информация", "Выберите пункт из списка")
                return

            item_text = self.sendto_listbox.get(selection[0])
            sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'SendTo')

            if "❌" in item_text:
                original_name = item_text[2:].replace(" (отключен)", "")
                item_path = os.path.join(sendto_path, f"{original_name}.DeskLink_no")
                enabled_path = os.path.join(sendto_path, original_name)

                try:
                    if os.path.exists(item_path) and not os.path.exists(enabled_path):
                        os.rename(item_path, enabled_path)
                        self.refresh_sendto_list()
                    else:
                        messagebox.showinfo("Информация", f"Пункт '{original_name}' уже включен или не найден")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось включить пункт: {str(e)}")
            else:
                if "✅" in item_text:
                    original_name = item_text[2:]
                else:
                    original_name = item_text

                item_path = os.path.join(sendto_path, original_name)
                disabled_path = os.path.join(sendto_path, f"{original_name}.DeskLink_no")

                try:
                    if os.path.exists(item_path) and not os.path.exists(disabled_path):
                        os.rename(item_path, disabled_path)
                        self.refresh_sendto_list()
                    else:
                        messagebox.showinfo("Информация", f"Пункт '{original_name}' уже отключен или не найден")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось отключить пункт: {str(e)}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта: {str(e)}")

    def refresh_sendto_list(self):
        """Обновление списка пунктов меню Отправить"""
        try:
            self.sendto_listbox.delete(0, tk.END)
            sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'SendTo')

            if not os.path.exists(sendto_path):
                self.sendto_listbox.insert(0, "Папка SendTo не найдена")
                return

            items = os.listdir(sendto_path)

            if not items:
                self.sendto_listbox.insert(0, "Папка SendTo пуста")
                return

            has_enabled_items = False
            has_disabled_items = False

            display_items = []

            for item in sorted(items):
                if item.lower() == "desktop.ini":
                    continue

                item_path = os.path.join(sendto_path, item)

                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if item.endswith('.DeskLink_no'):
                        original_name = item[:-12]
                        display_text = f"❌ {original_name} (отключен)"
                        has_disabled_items = True
                    else:
                        display_text = f"✅ {item}"
                        has_enabled_items = True

                    display_items.append((item, display_text))

            if has_enabled_items or (not has_enabled_items and not has_disabled_items):
                self.sendto_toggle_all_btn.config(text="Отключить всё")
            else:
                self.sendto_toggle_all_btn.config(text="Включить всё")

            for _, display_text in display_items:
                self.sendto_listbox.insert(tk.END, display_text)

        except Exception as e:
            self.sendto_listbox.insert(0, f"Ошибка: {e}")

    def update_menu_items_status(self):
        """Обновление статусов всех пунктов меню"""
        try:
            for btn_info in self.menu_item_buttons:
                item = btn_info["data"]

                exists = item["check_command"]()
                item["enabled"] = exists

                self.update_indicator(item, exists)

                btn = btn_info["button"]
                btn.configure(text=item['name'])

        except:
            pass

    def set_registry_key_permissions_simple(self, key_path):
        """Упрощенная установка прав доступа на ключ реестра"""
        try:
            reg_command = f'reg add "HKLM\\{key_path}" /f'

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["reg", "add", f"HKLM\\{key_path}", "/f"],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=10,
                startupinfo=startupinfo,
                shell=True
            )

            if result.returncode != 0:
                ps_command = f'''
                try {{
                    $key = "HKLM:\\{key_path}"
                    New-Item -Path $key -Force -ErrorAction Stop | Out-Null
                    return $true
                }} catch {{
                    return $false
                }}
                '''

                result = subprocess.run(
                    ["powershell", "-Command", ps_command],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    timeout=10,
                    startupinfo=startupinfo,
                    shell=True
                )

                return result.returncode == 0

            return True

        except:
            return False

    def take_ownership_registry_key_simple(self, key_path):
        """Упрощенное взятие владения ключом реестра"""
        try:
            reg_owner_command = f'reg add "HKLM\\{key_path}" /f /reg:64'

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["reg", "add", f"HKLM\\{key_path}", "/f", "/reg:64"],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=10,
                startupinfo=startupinfo,
                shell=True
            )

            if result.returncode != 0:
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
                    }} catch {{
                        return $false
                    }}
                }}
                '''

                result = subprocess.run(
                    ["powershell", "-Command", ps_command],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    timeout=10,
                    startupinfo=startupinfo,
                    shell=True
                )

                return result.returncode == 0

            return True

        except:
            return False

    def set_registry_key_permissions(self, hive, key_path):
        """Установка полных прав доступа на ключ реестра через subprocess"""
        try:
            if hive == winreg.HKEY_CLASSES_ROOT:
                root = "HKCR"
            elif hive == winreg.HKEY_CURRENT_USER:
                root = "HKCU"
            elif hive == winreg.HKEY_LOCAL_MACHINE:
                root = "HKLM"
            else:
                root = "HKCR"

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
            }} catch {{
                Write-Error "Ошибка: $_"
            }}
            '''

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=30,
                startupinfo=startupinfo
            )
            return result.returncode == 0

        except:
            return False

    def take_ownership_registry_key(self, hive, key_path):
        """Взятие владения ключом реестра через subprocess"""
        try:
            if hive == winreg.HKEY_CLASSES_ROOT:
                root = "HKCR"
            elif hive == winreg.HKEY_CURRENT_USER:
                root = "HKCU"
            elif hive == winreg.HKEY_LOCAL_MACHINE:
                root = "HKLM"
            else:
                root = "HKCR"

            full_path = f"{root}\\{key_path}"

            ps_command = f'''
            $key = "Registry::{full_path}"
            try {{
                $acl = Get-Acl $key
                $owner = [System.Security.Principal.NTAccount]"BUILTIN\\Administrators"
                $acl.SetOwner($owner)
                Set-Acl $key $acl
            }} catch {{
                Write-Error "Ошибка: $_"
            }}
            '''

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                encoding='cp866',
                timeout=30,
                startupinfo=startupinfo
            )

            return result.returncode == 0

        except:
            return False

    def rename_registry_key(self, hive, old_key_path, new_key_path):
        """Переименование ключа реестра с обработкой прав доступа"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                self.copy_registry_key(hive, old_key_path, new_key_path)
                self.delete_registry_key(hive, old_key_path)
                return True

            except PermissionError as pe:
                if self.take_ownership_registry_key_simple(old_key_path):
                    if self.set_registry_key_permissions_simple(old_key_path):
                        try:
                            self.copy_registry_key(hive, old_key_path, new_key_path)
                            self.delete_registry_key(hive, old_key_path)
                            return True
                        except PermissionError:
                            continue

            except FileNotFoundError as fnfe:
                try:
                    with winreg.OpenKey(hive, new_key_path):
                        return True
                except FileNotFoundError:
                    return False

            except:
                continue

        return False

    def copy_registry_key(self, hive, source_path, dest_path):
        """Копирование ключа реестра со всеми значениями и подключами"""
        try:
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
                            subkey_source_path = source_path + "\\" + subkey_name
                            subkey_dest_path = dest_path + "\\" + subkey_name
                            self.copy_registry_key(hive, subkey_source_path, subkey_dest_path)
                            i += 1
                        except OSError:
                            break

            return True

        except:
            raise

    def delete_registry_key(self, hive, key_path):
        """Удаление ключа реестра"""
        try:
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    i = 0
                    while True:
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            subkey_path = key_path + "\\" + subkey_name
                            self.delete_registry_key(hive, subkey_path)
                            i += 1
                        except OSError:
                            break
            except FileNotFoundError:
                return True

            winreg.DeleteKey(hive, key_path)
            return True

        except PermissionError as pe:
            raise
        except FileNotFoundError as fnfe:
            return True
        except:
            raise

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()