import hashlib
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import configparser
from datetime import datetime


class AboutModule:
    def __init__(self, parent, version, dll_info=None):
        self.parent = parent
        self.version = version
        self.main_app = self.get_main_app(parent)

        self.dll_info = dll_info
        if self.dll_info is None:
            try:
                from main import DLLChecker
                self.dll_info = DLLChecker.get_dll_info()
            except ImportError:
                self.dll_info = {
                    'expected_sha1': 'F62A4326A8300D9B828744E24314B98433510195',
                    'actual_hash': None,
                    'dll_path': None,
                    'hash_match': False,
                    'dll_found': False
                }

        self.mod_expected_hashes = {
            "imageres.dll": "DDD5D7462A26F7837CD02A7667615354F9377A0A727BF880D825676D6DDC4CB7",
            "imagesp1.dll": "466E64D491C7B3D7051FBDC0907CD813D7A6435D18BF4B00CE0C3191B80E0541",
            "networkexplorer.dll": "7BF68211CFBDDE2E377C992CEC2F5A0C1F41B1F12BE3A99733A549AB63A8C961",
            "DDORes.dll": "4802C8F0BAAE3765FDFF16F09F56647B1C8B20889E4C531B773442EC068CBC3B",
            "mmres.dll": "4651BB26196D0ED493A26F4E401C8FAC0465E7F9DE2ECFED1642155A6EAF3703"
        }

        self.frame = ttk.Frame(parent)
        self.create_widgets()
        self.show()

    def get_main_app(self, parent):
        try:
            root = parent
            while root.master:
                root = root.master
            if hasattr(root, 'main_app'):
                return root.main_app
        except Exception:
            pass
        return None

    def get_dll_status_text(self):
        if not self.dll_info['dll_found']:
            return "Не найден"
        if self.dll_info['hash_match']:
            return "Совпадает"
        return "Не совпадает"

    def get_dll_status_color(self):
        if not self.dll_info['dll_found']:
            return '#e74c3c'
        if self.dll_info['hash_match']:
            return '#27ae60'
        return '#e67e22'

    def create_widgets(self):
        main_container = ttk.Frame(self.frame)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        ttk.Label(main_container, text="О программе",
                  font=('Arial', 18, 'bold'), foreground='#2c3e50').pack(pady=1)
        ttk.Label(main_container, text=f"Interface Master версия {self.version}",
                  font=('Arial', 10), foreground='#3498db').pack(pady=(0, 15))

        info_frame = ttk.Frame(main_container)
        info_frame.pack(fill=tk.BOTH, expand=True)

        desc_frame = ttk.LabelFrame(info_frame, text="Описание", padding=10)
        desc_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(desc_frame,
                  text="Программа для настройки интерфейса Windows\nПомощник в кастомизации внешнего вида системы",
                  font=('Arial', 9), justify=tk.CENTER).pack()

        features_frame = ttk.LabelFrame(info_frame, text="Что можно настроить", padding=10)
        features_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(features_frame,
                  text="• Иконки рабочего стола          • Иконки дисков\n"
                       "• Стрелки ярлыков                    • Контекстное меню\n"
                       "• Панель задач                           • Проводник\n"
                       "                            • Модификации",
                  font=('Arial', 9), justify=tk.LEFT).pack()

        dll_frame = ttk.LabelFrame(info_frame, text="Файл imaster.dll", padding=10)
        dll_frame.pack(fill=tk.X, pady=(0, 15))
        self.create_dll_section(dll_frame)

        dev_frame = ttk.LabelFrame(info_frame, text="Разработка", padding=10)
        dev_frame.pack(fill=tk.X)
        ttk.Label(dev_frame,
                  text="Идея: EvKyz\nРеализация: DeepSeek AI\nСборка: 6 ноября 2026\nЛицензия: MIT License",
                  font=('Arial', 8), justify=tk.CENTER, foreground='#7f8c8d').pack()

        buttons_container = ttk.Frame(main_container)
        buttons_container.pack(fill=tk.X, pady=(10, 5))

        left_buttons_frame = ttk.Frame(buttons_container)
        left_buttons_frame.pack(side=tk.LEFT)

        save_btn = ttk.Button(left_buttons_frame, text="Сохранить настройки",
                              command=self.save_settings_to_ini, width=20)
        save_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.create_tooltip(save_btn, "Сохранить все текущие настройки в INI файл")

        load_btn = ttk.Button(left_buttons_frame, text="Загрузить настройки",
                              command=self.load_settings_from_ini, width=20)
        load_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.create_tooltip(load_btn, "Загрузить настройки из INI файла и применить их")

        reset_btn = ttk.Button(buttons_container, text="Сбросить все настройки",
                               command=self.reset_all_settings, width=25)
        reset_btn.pack(side=tk.RIGHT)
        self.create_tooltip(reset_btn, "Сбросить все настройки программы к стандартным значениям Windows")

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

    def create_dll_section(self, dll_frame):
        container = tk.Frame(dll_frame, bg='#f5f5f5', padx=5, pady=5)
        container.pack(fill=tk.BOTH, expand=True)

        tk.Label(container, text=f"Статус: {self.get_dll_status_text()}",
                 font=('Consolas', 9, 'bold'), fg=self.get_dll_status_color(),
                 bg='#f5f5f5', anchor='w').pack(fill=tk.X, pady=(0, 5))

        if self.dll_info['dll_found'] and self.dll_info['dll_path']:
            tk.Label(container, text=f"Путь: {self.dll_info['dll_path']}",
                     font=('Consolas', 8), bg='#f5f5f5', anchor='w').pack(fill=tk.X, pady=(0, 3))
            if self.dll_info['actual_hash']:
                tk.Label(container, text=f"Хеш SHA1: {self.dll_info['actual_hash']}",
                         font=('Consolas', 8), bg='#f5f5f5', anchor='w').pack(fill=tk.X)
        else:
            tk.Label(container, text="Файл imaster.dll не найден",
                     font=('Consolas', 8), bg='#f5f5f5', anchor='w').pack(fill=tk.X)

    def log_to_file(self, message, log_name="apply_settings.log"):
        try:
            log_file_path = os.path.join(os.path.expanduser("~"), "Desktop", log_name)
            os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {message}\n")
        except Exception:
            pass

    def get_registry_value(self, key_path, value_name, hive='HKCU'):
        try:
            cmd = ['reg', 'query', f'{hive}\\{key_path}', '/v', value_name]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='cp866',
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if value_name in line:
                        parts = line.split()
                        if len(parts) >= 3:
                            return parts[-1]
        except Exception:
            pass
        return None

    def collect_all_settings(self):
        return {
            'metadata': {
                'version': self.version,
                'export_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'program': 'Interface Master'
            },
            'desktop': self.collect_desktop_settings(),
            'disk': self.collect_disk_settings(),
            'str': self.collect_shortcut_settings(),
            'contextmenu': self.collect_contextmenu_settings(),
            'taskbar': self.collect_taskbar_settings(),
            'explorer': self.collect_explorer_settings()
        }

    def collect_desktop_settings(self):
        settings = {}

        items = {
            'user': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon',
            'computer': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon',
            'network': r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon'
        }
        for item, path in items.items():
            val = self.get_registry_value(path, '', 'HKCU')
            if val:
                settings[f'icon_{item}'] = val

        recycle_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon'
        empty_val = self.get_registry_value(recycle_path, '', 'HKCU')

        if empty_val:
            if ',-' in empty_val:
                idx = empty_val.split(',-')[-1]
                recycle_pairs = {
                    '4': ('4', '5'), '5': ('4', '5'),
                    '24': ('24', '25'), '25': ('24', '25'),
                    '34': ('34', '35'), '35': ('34', '35'),
                    '37': ('37', '38'), '38': ('37', '38'),
                    '44': ('44', '45'), '45': ('44', '45'),
                    '64': ('64', '65'), '65': ('64', '65'),
                    '74': ('74', '75'), '75': ('74', '75'),
                    '84': ('84', '85'), '85': ('84', '85')
                }
                if idx in recycle_pairs:
                    empty_idx, full_idx = recycle_pairs[idx]
                    dll_path = empty_val.split(',-')[0]
                    settings['icon_recycle_bin_empty'] = f"{dll_path},-{empty_idx}"
                    settings['icon_recycle_bin_full'] = f"{dll_path},-{full_idx}"
                else:
                    settings['icon_recycle_bin_empty'] = empty_val
                    full_val = self.get_registry_value(recycle_path, 'Full', 'HKCU')
                    settings['icon_recycle_bin_full'] = full_val or empty_val
            else:
                settings['icon_recycle_bin_empty'] = empty_val
                full_val = self.get_registry_value(recycle_path, 'Full', 'HKCU')
                settings['icon_recycle_bin_full'] = full_val or empty_val

        libraries_val = self.get_registry_value(
            r'CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon', '', 'HKCR')
        if libraries_val:
            settings['icon_libraries'] = libraries_val

        control_panel_clsids = [
            '{26EE0668-A00A-44D7-9371-BEB064C98683}',
            '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}',
            '{21EC2020-3AEA-1069-A2DD-08002B30309D}'
        ]
        for clsid in control_panel_clsids:
            path = rf'Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{clsid}\DefaultIcon'
            val = self.get_registry_value(path, '', 'HKCU')
            if val:
                settings[f'icon_control_panel_{clsid}'] = val

        return settings

    def collect_disk_settings(self):
        settings = {}
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons'],
                capture_output=True, text=True, encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if 'HKEY_LOCAL_MACHINE' in line:
                        drive = line.split('\\')[-1].strip()
                        if drive and len(drive) == 1 and drive.isalpha():
                            icon_path = f'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons\\{drive}\\DefaultIcon'
                            result2 = subprocess.run(['reg', 'query', icon_path, '/ve'],
                                                     capture_output=True, text=True, encoding='cp866',
                                                     creationflags=subprocess.CREATE_NO_WINDOW)
                            if result2.returncode == 0:
                                for line2 in result2.stdout.split('\n'):
                                    if 'REG_' in line2:
                                        parts = line2.split()
                                        if len(parts) >= 3:
                                            settings[f'icon_{drive}'] = parts[-1]
        except Exception:
            pass
        return settings

    def collect_shortcut_settings(self):
        settings = {}
        arrow_val = self.get_registry_value(
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons', '29', 'HKLM')
        if arrow_val is None:
            settings['arrow_style'] = 'default'
        elif arrow_val == '':
            settings['arrow_style'] = 'none'
        elif 'shell32.dll,-30' in str(arrow_val):
            settings['arrow_style'] = 'small'
        else:
            settings['arrow_style'] = 'custom'

        template_val = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\NamingTemplates',
            'ShortcutNameTemplate', 'HKCU')
        if template_val:
            if '"%s — ярлык.lnk"' in str(template_val):
                settings['shortcut_text'] = 'default'
            elif '"%s.lnk"' in str(template_val):
                settings['shortcut_text'] = 'removed'
            else:
                settings['shortcut_text'] = 'custom'
        else:
            settings['shortcut_text'] = 'default'
        return settings

    def collect_contextmenu_settings(self):
        settings = {}
        items = {
            'copy_as_path': r'HKLM\SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu',
            'pin_to_start': r'HKLM\SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen',
            'pin_to_taskbar': r'HKLM\SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}',
            'take_ownership': r'HKCR\*\shell\runas',
            'copy_move': r'HKLM\SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
            'admin_menu': r'HKCU\Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin',
            'system_menu': r'HKCU\Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System',
            'task_manager': r'HKCR\DesktopBackground\Shell\TaskManager',
            'restart_explorer': r'HKCR\DesktopBackground\Shell\RestartExplorer',
            'clipboard_clear': r'HKCR\DesktopBackground\Shell\Clip',
            'regedit': r'HKCR\DesktopBackground\Shell\Regedit',
            'power_menu': r'HKCR\DesktopBackground\Shell\Shutdown'
        }
        for name, path in items.items():
            hive = path.split('\\')[0]
            key_path = '\\'.join(path.split('\\')[1:])
            try:
                result = subprocess.run(['reg', 'query', f'{hive}\\{key_path}'],
                                        capture_output=True,
                                        creationflags=subprocess.CREATE_NO_WINDOW)
                settings[name] = result.returncode == 0
            except Exception:
                settings[name] = False
        return settings

    def collect_taskbar_settings(self):
        settings = {}
        short_date = self.get_registry_value(r'Control Panel\International', 'sShortDate', 'HKCU')
        settings['show_weekday'] = 'ddd' in str(short_date) if short_date else False

        notif = self.get_registry_value(r'SOFTWARE\Policies\Microsoft\Windows\Explorer',
                                        'DisableNotificationCenter', 'HKLM')
        settings['notification_center'] = notif != '0x1'

        small = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'TaskbarSmallIcons', 'HKCU')
        settings['small_icons'] = small == '0x1'

        search = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Search', 'SearchboxTaskbarMode', 'HKCU')
        if search == '0x0':
            settings['search_mode'] = 0
        elif search == '0x1':
            settings['search_mode'] = 1
        else:
            settings['search_mode'] = 2

        hide_people = self.get_registry_value(
            r'Software\Policies\Microsoft\Windows\Explorer', 'HidePeopleBar', 'HKCU')
        settings['show_people'] = hide_people != '0x1'

        task_view = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'ShowTaskViewButton', 'HKCU')
        settings['show_task_view'] = task_view != '0x0'

        glom = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'TaskbarGlomLevel', 'HKCU')
        if glom == '0x0':
            settings['grouping_mode'] = 0
        elif glom == '0x1':
            settings['grouping_mode'] = 1
        elif glom == '0x2':
            settings['grouping_mode'] = 2
        else:
            settings['grouping_mode'] = 0

        meet = self.get_registry_value(
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer',
            'HideSCAMeetNow', 'HKLM')
        settings['show_meet_now'] = meet != '0x1'

        return settings

    def collect_explorer_settings(self):
        settings = {}
        bitlocker = self.check_registry_key_exists(r'Drive\shell\encrypt-bde\ProgrammaticAccessOnly', 'HKCR')
        settings['bitlocker_enabled'] = not bitlocker

        photo = self.get_registry_value(
            r'SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations', '.jpg', 'HKLM')
        settings['photo_viewer_enabled'] = photo is not None

        hide_ext = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced',
            'HideFileExt', 'HKCU')
        settings['show_file_extensions'] = hide_ext == '0x0'

        no_recent = self.get_registry_value(
            r'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer',
            'NoRecentDocsHistory', 'HKCU')
        settings['recent_folders_enabled'] = no_recent != '0x1'

        return settings

    def check_registry_key_exists(self, key_path, hive='HKLM'):
        try:
            result = subprocess.run(['reg', 'query', f'{hive}\\{key_path}'],
                                    capture_output=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            return result.returncode == 0
        except Exception:
            return False

    def save_settings_to_ini(self):
        try:
            file_path = filedialog.asksaveasfilename(
                title="Сохранить настройки",
                defaultextension=".ini",
                filetypes=[("INI files", "*.ini")],
                initialfile=f"interface_master_{datetime.now().strftime('%Y%m%d_%H%M%S')}.ini")
            if not file_path:
                return

            config = configparser.ConfigParser()
            config.optionxform = str

            for section, values in self.collect_all_settings().items():
                config[section] = {}
                for key, value in values.items():
                    value_str = str(value)
                    if '%' in value_str:
                        value_str = value_str.replace('%', '%%')
                    config[section][key] = value_str

            with open(file_path, 'w', encoding='utf-8') as f:
                config.write(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{str(e)}")

    def load_settings_from_ini(self):
        try:
            file_path = filedialog.askopenfilename(title="Загрузить настройки",
                                                   filetypes=[("INI files", "*.ini")])
            if not file_path:
                return

            config = configparser.ConfigParser()
            config.optionxform = str
            config.read(file_path, encoding='utf-8')

            self.log_to_file("=" * 60, "apply_settings.log")
            self.log_to_file("ЗАГРУЗКА НАСТРОЕК ИЗ INI", "apply_settings.log")
            if 'str' in config:
                self.log_to_file("Секция [str]:", "apply_settings.log")
                for key, value in config['str'].items():
                    self.log_to_file(f"  {key} = {value}", "apply_settings.log")
            self.log_to_file("-" * 60, "apply_settings.log")

            self.apply_settings_from_config(config)
            self.restart_explorer_silent()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить настройки:\n{str(e)}")

    def apply_settings_from_config(self, config):
        if 'str' in config:
            self._apply_str_settings(config)
        if 'taskbar' in config:
            self._apply_taskbar_settings(config)
        if 'explorer' in config:
            self._apply_explorer_settings(config)
        if 'desktop' in config:
            self._apply_desktop_icons(config)
        if 'disk' in config:
            self._apply_disk_icons(config)
        if 'contextmenu' in config:
            self._apply_contextmenu_settings(config)

        self.log_to_file("=" * 60, "apply_settings.log")
        self.log_to_file("ПРИМЕНЕНИЕ НАСТРОЕК ЗАВЕРШЕНО", "apply_settings.log")
        self.log_to_file("=" * 60, "apply_settings.log")

    def _apply_str_settings(self, config):
        self.log_to_file("=== ПРИМЕНЕНИЕ НАСТРОЕК ЯРЛЫКОВ ===", "apply_settings.log")
        try:
            from str import Module as StrModule
            temp = ttk.Frame(self.parent)
            s = StrModule(temp)

            arrow_changed = False
            text_changed = False

            if 'arrow_style' in config['str']:
                style = config['str']['arrow_style']
                self.log_to_file(f"  arrow_style = {style}", "apply_settings.log")
                if style == 'default':
                    s.arrow_choice.set(1)
                    arrow_changed = True
                elif style == 'small':
                    s.arrow_choice.set(2)
                    arrow_changed = True
                elif style == 'none':
                    s.arrow_choice.set(3)
                    arrow_changed = True
                else:
                    self.log_to_file(f"    Неизвестный стиль '{style}'", "apply_settings.log")

            if 'shortcut_text' in config['str']:
                txt = config['str']['shortcut_text']
                self.log_to_file(f"  shortcut_text = {txt}", "apply_settings.log")
                if txt == 'default':
                    s.shortcut_text_choice.set(1)
                    text_changed = True
                else:
                    s.shortcut_text_choice.set(2)
                    text_changed = True

            if arrow_changed or text_changed:
                s.apply_changes()
                self.log_to_file(
                    f"    Изменения применены (стрелки: {arrow_changed}, текст: {text_changed})",
                    "apply_settings.log")

            s.hide()
            temp.destroy()
        except Exception as e:
            self.log_to_file(f"  ОШИБКА str: {e}", "apply_settings.log")

    def _apply_taskbar_settings(self, config):
        self.log_to_file("=== ПРИМЕНЕНИЕ НАСТРОЕК ПАНЕЛИ ЗАДАЧ ===", "apply_settings.log")
        try:
            from taskbar import Module as TaskbarModule
            temp = ttk.Frame(self.parent)
            t = TaskbarModule(temp)

            if 'show_weekday' in config['taskbar']:
                current = t.check_weekday_state()
                target = config['taskbar']['show_weekday'].lower() == 'true'
                if current != target:
                    t.toggle_weekday()
                    self.log_to_file(f"  День недели: {current} -> {target}", "apply_settings.log")
            if 'notification_center' in config['taskbar']:
                current = not t.check_notification_center_state()
                target = config['taskbar']['notification_center'].lower() == 'true'
                if current != target:
                    t.toggle_notification_center()
            if 'small_icons' in config['taskbar']:
                current = t.check_taskbar_size_state()
                target = config['taskbar']['small_icons'].lower() == 'true'
                if current != target:
                    t.toggle_taskbar_size()
            if 'search_mode' in config['taskbar']:
                target = int(config['taskbar']['search_mode'])
                current = t.check_search_state()
                if current != target:
                    for _ in range(10):
                        if t.check_search_state() == target:
                            break
                        t.toggle_search()
            if 'show_people' in config['taskbar']:
                current = not t.check_people_state()
                target = config['taskbar']['show_people'].lower() == 'true'
                if current != target:
                    t.toggle_people()
            if 'show_task_view' in config['taskbar']:
                current = not t.check_task_view_state()
                target = config['taskbar']['show_task_view'].lower() == 'true'
                if current != target:
                    t.toggle_task_view()
            if 'grouping_mode' in config['taskbar']:
                target = int(config['taskbar']['grouping_mode'])
                current = t.check_grouping_state()
                if current != target:
                    for _ in range(10):
                        if t.check_grouping_state() == target:
                            break
                        t.toggle_grouping()
            if 'show_meet_now' in config['taskbar'] and t.is_windows_version_supported():
                current = not t.check_meet_now_state()
                target = config['taskbar']['show_meet_now'].lower() == 'true'
                if current != target:
                    t.toggle_meet_now()

            t.hide()
            temp.destroy()
        except Exception as e:
            self.log_to_file(f"  ОШИБКА taskbar: {e}", "apply_settings.log")

    def _apply_explorer_settings(self, config):
        self.log_to_file("=== ПРИМЕНЕНИЕ НАСТРОЕК ПРОВОДНИКА ===", "apply_settings.log")
        try:
            from explorer import ExplorerModule
            temp = ttk.Frame(self.parent)
            e = ExplorerModule(temp)

            if 'bitlocker_enabled' in config['explorer']:
                target = config['explorer']['bitlocker_enabled'].lower() == 'true'
                if target != e.bitlocker_enabled:
                    e.toggle_bitlocker()
            if 'photo_viewer_enabled' in config['explorer']:
                target = config['explorer']['photo_viewer_enabled'].lower() == 'true'
                if target != e.photo_viewer_enabled:
                    e.toggle_photo_viewer()
            if 'show_file_extensions' in config['explorer']:
                target = config['explorer']['show_file_extensions'].lower() == 'true'
                if target != e.file_ext_enabled:
                    e.toggle_file_ext()
            if 'recent_folders_enabled' in config['explorer']:
                target = config['explorer']['recent_folders_enabled'].lower() == 'true'
                if target != e.recent_folders_enabled:
                    e.toggle_recent_folders()

            e.hide()
            temp.destroy()
        except Exception as e:
            self.log_to_file(f"  ОШИБКА explorer: {e}", "apply_settings.log")

    def _apply_desktop_icons(self, config):
        self.log_to_file("=== ПРИМЕНЕНИЕ НАСТРОЕК ИКОНОК РАБОЧЕГО СТОЛА ===", "apply_settings.log")
        recycle_path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon"

        for key, value in config['desktop'].items():
            if not key.startswith('icon_'):
                continue
            icon_type = key.replace('icon_', '')
            self.log_to_file(f"  {icon_type} = {value}", "apply_settings.log")

            if icon_type == 'user':
                path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon"
                subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)
            elif icon_type == 'computer':
                path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon"
                subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)
            elif icon_type == 'network':
                path = r"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon"
                subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)
            elif icon_type == 'recycle_bin_empty':
                subprocess.run(['reg', 'add', recycle_path, '/ve', '/d', str(value), '/f'], capture_output=True)
            elif icon_type == 'recycle_bin_full':
                subprocess.run(['reg', 'add', recycle_path, '/v', 'Full', '/d', str(value), '/f'], capture_output=True)
            elif icon_type == 'libraries':
                path = r"HKCR\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"
                subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)
            elif icon_type.startswith('control_panel'):
                if icon_type == 'control_panel':
                    clsid_list = [
                        '{26EE0668-A00A-44D7-9371-BEB064C98683}',
                        '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}',
                        '{21EC2020-3AEA-1069-A2DD-08002B30309D}'
                    ]
                    for clsid in clsid_list:
                        path = rf"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{clsid}\DefaultIcon"
                        subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)
                else:
                    clsid = icon_type.replace('control_panel_', '')
                    path = rf"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{clsid}\DefaultIcon"
                    subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)

        subprocess.run(['ie4uinit.exe', '-show'], capture_output=True)
        self.log_to_file("  Иконки рабочего стола применены", "apply_settings.log")

    def _apply_disk_icons(self, config):
        self.log_to_file("=== ПРИМЕНЕНИЕ НАСТРОЕК ИКОНОК ДИСКОВ ===", "apply_settings.log")
        subprocess.run(
            ['reg', 'delete', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons', '/f'],
            capture_output=True)
        for key, value in config['disk'].items():
            if key.startswith('icon_') and value:
                drive = key.replace('icon_', '').upper()
                if len(drive) == 1 and drive.isalpha():
                    path = f"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons\\{drive}\\DefaultIcon"
                    subprocess.run(['reg', 'add', path, '/ve', '/d', str(value), '/f'], capture_output=True)
                    self.log_to_file(f"  Диск {drive}: {value}", "apply_settings.log")

    def _apply_contextmenu_settings(self, config):
        self.log_to_file("=== ПРИМЕНЕНИЕ НАСТРОЕК КОНТЕКСТНОГО МЕНЮ ===", "apply_settings.log")
        try:
            from contextmenu import ContextMenuModule
            temp = ttk.Frame(self.parent)
            c = ContextMenuModule(temp)

            menu_mapping = {
                'copy_as_path': 'Копировать как путь',
                'pin_to_start': 'Закрепить на начальном экране',
                'pin_to_taskbar': 'Закрепить на Панели задач',
                'take_ownership': 'Стать владельцем',
                'copy_move': 'Копировать/Переместить в папку',
                'admin_menu': 'Администрирование',
                'system_menu': 'Система',
                'task_manager': 'Диспетчер задач',
                'restart_explorer': 'Перезапустить explorer',
                'clipboard_clear': 'Очистить буфер',
                'regedit': 'Редактор реестра',
                'power_menu': 'Выключение/Перезагрузка/Гибернация'
            }

            if hasattr(c, 'desktop_menu_tab'):
                dt = c.desktop_menu_tab
                for key, target_state in config['contextmenu'].items():
                    target = target_state.lower() == 'true'
                    menu_name = menu_mapping.get(key, key)
                    for btn_info in dt.menu_item_buttons:
                        item = btn_info['data']
                        if item['name'] == menu_name:
                            current = item['check_command']()
                            if current != target:
                                if target:
                                    item['add_command']()
                                else:
                                    item['remove_command']()
                            break

            c.hide()
            temp.destroy()
        except Exception as e:
            self.log_to_file(f"  ОШИБКА contextmenu: {e}", "apply_settings.log")

    def restart_explorer_silent(self):
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'], capture_output=True)
            import time
            time.sleep(0.5)
            subprocess.Popen(['explorer.exe'])
        except Exception:
            pass

    def check_modified_files(self):
        modified = []
        for fname, expected in self.mod_expected_hashes.items():
            path = os.path.join(os.environ["WINDIR"], "System32", fname)
            if os.path.exists(path):
                try:
                    with open(path, "rb") as f:
                        h = hashlib.sha256(f.read()).hexdigest().upper()
                    if h == expected.upper():
                        modified.append(fname)
                except Exception:
                    pass
        return modified

    def restore_modified_files(self):
        restored = []
        for fname in self.mod_expected_hashes:
            path = os.path.join(os.environ["WINDIR"], "System32", fname)
            backup = f"{path}_orig"
            if os.path.exists(path):
                try:
                    with open(path, "rb") as f:
                        if hashlib.sha256(f.read()).hexdigest().upper() == self.mod_expected_hashes[fname].upper():
                            subprocess.run("taskkill /f /im explorer.exe", shell=True, capture_output=True)
                            os.remove(path)
                            if os.path.exists(backup):
                                os.rename(backup, path)
                                restored.append(fname)
                except Exception:
                    pass
        return restored

    def reset_all_settings(self):
        modified = self.check_modified_files()
        msg = "ВНИМАНИЕ! Это действие сбросит все настройки к стандартным значениям Windows.\n\n"
        msg += "Будут сброшены:\n• Иконки рабочего стола\n• Иконки дисков\n• Стрелки ярлыков\n"
        msg += "• Текст '- ярлык'\n• Пункты контекстного меню\n• Настройки Проводника\n• Настройки Панели задач\n"
        if modified:
            msg += "\n⚠️ БУДУТ ВОССТАНОВЛЕНЫ МОДИФИЦИРОВАННЫЕ ФАЙЛЫ:\n"
            for f in modified:
                msg += f"  • {f}\n"
        msg += "\n\nПродолжить?"

        if not messagebox.askyesno("Сброс всех настроек", msg, icon='warning'):
            return

        errors = []

        if modified:
            self.restore_modified_files()

        try:
            from desktop import Module as DesktopModule
            temp = ttk.Frame(self.parent)
            d = DesktopModule(temp, module_name="desktop")
            d.reset_icons()
            d.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"desktop: {e}")

        try:
            subprocess.run(
                ['reg', 'delete', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons', '/f'],
                capture_output=True)
        except Exception as e:
            errors.append(f"disk: {e}")

        try:
            from str import Module as StrModule
            temp = ttk.Frame(self.parent)
            s = StrModule(temp)
            s.arrow_choice.set(1)
            s.shortcut_text_choice.set(1)
            s.apply_changes()
            s.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"str: {e}")

        try:
            from contextmenu import ContextMenuModule
            temp = ttk.Frame(self.parent)
            c = ContextMenuModule(temp)
            if hasattr(c, 'desktop_menu_tab'):
                c.desktop_menu_tab.set_default_settings()
            c.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"contextmenu: {e}")

        try:
            from explorer import ExplorerModule
            temp = ttk.Frame(self.parent)
            e = ExplorerModule(temp)
            for guid in e.folders:
                e.checkbox_vars[guid].set(True)
            e.apply_folder_changes()
            if not e.bitlocker_enabled:
                e.toggle_bitlocker()
            if not e.photo_viewer_enabled:
                e.toggle_photo_viewer()
            if e.file_ext_enabled:
                e.toggle_file_ext()
            e.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"explorer: {e}")

        try:
            from taskbar import Module as TaskbarModule
            temp = ttk.Frame(self.parent)
            t = TaskbarModule(temp)
            t.reset_all_settings()
            t.hide()
            temp.destroy()
        except Exception as e:
            errors.append(f"taskbar: {e}")

        self.restart_explorer_silent()

        if errors:
            messagebox.showwarning("Сброс завершён с ошибками", "\n".join(errors[:5]))
        else:
            messagebox.showinfo("Успех", "Все настройки успешно сброшены!")

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.frame.pack_forget()