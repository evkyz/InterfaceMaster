import subprocess
import sys
import re
import time
import tkinter as tk
from tkinter import ttk, messagebox

_CREATE_NO_WINDOW = 0
if sys.platform == "win32":
    _CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW


def _run(args, **kwargs):
    kwargs.setdefault("creationflags", _CREATE_NO_WINDOW)
    return subprocess.run(args, **kwargs)


def _popen(args, **kwargs):
    kwargs.setdefault("creationflags", _CREATE_NO_WINDOW)
    return subprocess.Popen(args, **kwargs)


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.windows_version = self.get_windows_version()
        self.create_widgets()
        self.show()

    def get_windows_version(self):
        try:
            result = _run(['wmic', 'os', 'get', 'Version', '/value'],
                          capture_output=True, text=True, encoding='cp866')
            match = re.search(r'Version=(\d+\.\d+\.\d+)', result.stdout)
            if match:
                return match.group(1)

            result = _run(['reg', 'query',
                           'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion',
                           '/v', 'CurrentBuildNumber'],
                          capture_output=True, text=True, encoding='cp866')
            match = re.search(r'CurrentBuildNumber\s+REG_SZ\s+(\d+)', result.stdout)
            if match:
                return f"10.0.{match.group(1)}"
            return "10.0.0"
        except Exception:
            return "10.0.0"

    def is_windows_version_supported(self):
        try:
            parts = self.windows_version.split('.')
            if len(parts) >= 3:
                return int(parts[2]) >= 19044
        except Exception:
            pass
        return False

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

    def check_weekday_state(self):
        try:
            result = _run(['reg', 'query', 'HKCU\\Control Panel\\International',
                           '/v', 'sShortDate'],
                          capture_output=True, text=True, encoding='cp866')
            return "ddd" in result.stdout
        except Exception:
            return False

    def check_notification_center_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                           '/v', 'DisableNotificationCenter'],
                          capture_output=True, text=True, encoding='cp866')
            return "0x1" in result.stdout
        except Exception:
            return False

    def check_taskbar_size_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                           '/v', 'TaskbarSmallIcons'],
                          capture_output=True, text=True, encoding='cp866')
            return "0x1" in result.stdout
        except Exception:
            return False

    def check_search_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
                           '/v', 'SearchboxTaskbarMode'],
                          capture_output=True, text=True, encoding='cp866')
            if "0x0" in result.stdout:
                return 0
            elif "0x1" in result.stdout:
                return 1
            return 2
        except Exception:
            return 2

    def check_people_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                           '/v', 'HidePeopleBar'],
                          capture_output=True, text=True, encoding='cp866')
            return "0x1" in result.stdout
        except Exception:
            return False

    def check_task_view_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                           '/v', 'ShowTaskViewButton'],
                          capture_output=True, text=True, encoding='cp866')
            return "0x0" in result.stdout
        except Exception:
            return False

    def check_meet_now_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                           '/v', 'HideSCAMeetNow'],
                          capture_output=True, text=True, encoding='cp866')
            return "0x1" in result.stdout
        except Exception:
            return False

    def check_grouping_state(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                           '/v', 'TaskbarGlomLevel'],
                          capture_output=True, text=True, encoding='cp866')
            if "0x0" in result.stdout:
                return 0
            elif "0x1" in result.stdout:
                return 1
            elif "0x2" in result.stdout:
                return 2
            return 0
        except Exception:
            return 0

    def update_button_text(self, button, base_text, state, is_search=False,
                           is_grouping=False):
        if button == self.toggle_taskbar_size_button:
            button.config(text=f"{base_text} ({'Маленькие' if state else 'Большие'})")
        elif is_search:
            labels = {0: "Скрыто", 1: "Значок", 2: "Поле"}
            button.config(text=f"{base_text} ({labels.get(state, 'Поле')})")
        elif is_grouping:
            labels = {0: "Всегда", 1: "При заполнении", 2: "Никогда"}
            button.config(text=f"{base_text} ({labels.get(state, 'Всегда')})")
        else:
            button.config(text=f"{base_text} ({'Показано' if state else 'Скрыто'})")

    def create_widgets(self):
        ttk.Label(self.frame, text="Настройка Панели задач",
                  font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(pady=5)
        ttk.Label(self.frame,
                  text="Настройте внешний вид и функциональность Панели задач Windows",
                  font=('Arial', 9), foreground='#7f8c8d').pack(pady=(0, 10))

        settings_frame = ttk.LabelFrame(self.frame, text="Доступные настройки",
                                        padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        grouping_state = self.check_grouping_state()
        grouping_texts = {0: " (Всегда)", 1: " (При заполнении)", 2: " (Никогда)"}
        self.toggle_grouping_button = ttk.Button(
            settings_frame,
            text="Изменить группировку кнопок" + grouping_texts[grouping_state],
            command=self.toggle_grouping_with_update, width=40)
        self.toggle_grouping_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_grouping_button,
                            "Всегда → При заполнении → Никогда")

        size_state = self.check_taskbar_size_state()
        self.toggle_taskbar_size_button = ttk.Button(
            settings_frame,
            text="Переключить размер значков панели задач"
                 + (" (Маленькие)" if size_state else " (Большие)"),
            command=self.toggle_taskbar_size_with_update, width=40)
        self.toggle_taskbar_size_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_taskbar_size_button,
                            "Переключает между большими и маленькими значками на панели задач")

        search_state = self.check_search_state()
        search_texts = {0: " (Скрыто)", 1: " (Значок)", 2: " (Поле)"}
        self.toggle_search_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Поиск" + search_texts[search_state],
            command=self.toggle_search_with_update, width=40)
        self.toggle_search_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_search_button, "Поле → Значок → Скрыто")

        task_view_state = self.check_task_view_state()
        self.toggle_task_view_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Просмотр задач"
                 + (" (Показано)" if not task_view_state else " (Скрыто)"),
            command=self.toggle_task_view_with_update, width=40)
        self.toggle_task_view_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_task_view_button,
                            "Скрывает или показывает кнопку 'Просмотр задач' на панели задач")

        people_state = self.check_people_state()
        self.toggle_people_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Люди"
                 + (" (Показано)" if not people_state else " (Скрыто)"),
            command=self.toggle_people_with_update, width=40)
        self.toggle_people_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_people_button,
                            "Скрывает или показывает кнопку 'Люди' на панели задач")

        notification_state = self.check_notification_center_state()
        self.toggle_notification_center_button = ttk.Button(
            settings_frame,
            text="Добавить/Убрать Центр уведомлений"
                 + (" (Показано)" if not notification_state else " (Скрыто)"),
            command=self.toggle_notification_center_with_update, width=40)
        self.toggle_notification_center_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_notification_center_button,
                            "Включает или отключает Центр уведомлений Windows")

        meet_now_state = self.check_meet_now_state() if self.is_windows_version_supported() else False
        self.toggle_meet_now_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Провести собрание"
                 + (" (Показано)" if not meet_now_state else " (Скрыто)"),
            command=self.toggle_meet_now_with_update, width=40)
        self.toggle_meet_now_button.pack(pady=5, fill=tk.X)

        if self.is_windows_version_supported():
            self.create_tooltip(self.toggle_meet_now_button,
                                "Скрывает или показывает кнопку 'Провести собрание' на панели задач")
        else:
            self.toggle_meet_now_button.config(state='disabled')
            self.create_tooltip(self.toggle_meet_now_button,
                                "Эта функция доступна только в Windows 10 21H2 и выше")

        weekday_state = self.check_weekday_state()
        self.toggle_weekday_button = ttk.Button(
            settings_frame,
            text="Добавить/Убрать день недели в дате"
                 + (" (Показано)" if weekday_state else " (Скрыто)"),
            command=self.toggle_weekday_with_update, width=40)
        self.toggle_weekday_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_weekday_button,
                            "Добавляет или убирает отображение дня недели в системной дате")

        extra_frame = ttk.Frame(settings_frame)
        extra_frame.pack(fill=tk.X, pady=10)

        restart_btn = ttk.Button(extra_frame, text="Перезапустить Проводник",
                                 command=self.restart_explorer, width=18)
        restart_btn.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        self.create_tooltip(restart_btn,
                            "Перезапускает Проводник Windows для применения изменений")

        reset_btn = ttk.Button(extra_frame, text="Сбросить все настройки",
                               command=self.reset_all_settings, width=18)
        reset_btn.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        self.create_tooltip(reset_btn,
                            "Сбрасывает все настройки панели задач к значениям по умолчанию")

    def toggle_weekday_with_update(self):
        self.toggle_weekday()
        self.update_button_text(self.toggle_weekday_button,
                                "Добавить/Убрать день недели в дате",
                                self.check_weekday_state())

    def toggle_notification_center_with_update(self):
        self.toggle_notification_center()
        self.update_button_text(self.toggle_notification_center_button,
                                "Добавить/Убрать Центр уведомлений",
                                not self.check_notification_center_state())

    def toggle_taskbar_size_with_update(self):
        self.toggle_taskbar_size()
        self.update_button_text(self.toggle_taskbar_size_button,
                                "Переключить размер значков панели задач",
                                self.check_taskbar_size_state())

    def toggle_search_with_update(self):
        self.toggle_search()
        self.update_button_text(self.toggle_search_button, "Скрыть/Показать Поиск",
                                self.check_search_state(), is_search=True)

    def toggle_people_with_update(self):
        self.toggle_people()
        self.update_button_text(self.toggle_people_button, "Скрыть/Показать Люди",
                                not self.check_people_state())

    def toggle_task_view_with_update(self):
        self.toggle_task_view()
        self.update_button_text(self.toggle_task_view_button,
                                "Скрыть/Показать Просмотр задач",
                                not self.check_task_view_state())

    def toggle_meet_now_with_update(self):
        self.toggle_meet_now()
        self.update_button_text(self.toggle_meet_now_button,
                                "Скрыть/Показать Провести собрание",
                                not self.check_meet_now_state())

    def toggle_grouping_with_update(self):
        self.toggle_grouping()
        self.update_button_text(self.toggle_grouping_button,
                                "Изменить группировку кнопок",
                                self.check_grouping_state(), is_grouping=True)

    def toggle_weekday(self):
        try:
            result = _run(['reg', 'query', 'HKCU\\Control Panel\\International',
                           '/v', 'sShortDate'],
                          capture_output=True, text=True, encoding='cp866')
            if "ddd" in result.stdout:
                _run(['reg', 'add', 'HKCU\\Control Panel\\International',
                      '/V', 'sShortDate', '/D', 'dd.MM.yyyy', '/T', 'REG_SZ', '/f'],
                     check=True, capture_output=True, text=True, encoding='cp866')
            else:
                _run(['reg', 'add', 'HKCU\\Control Panel\\International',
                      '/V', 'sShortDate', '/D', 'ddd dd.MM.yyyy', '/T', 'REG_SZ', '/f'],
                     check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_notification_center(self):
        try:
            try:
                result = _run(['reg', 'query',
                               'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                               '/v', 'DisableNotificationCenter'],
                              capture_output=True, text=True, encoding='cp866')
                current_value = 1 if "0x1" in result.stdout else 0
            except Exception:
                current_value = 0

            if current_value == 1:
                try:
                    _run(['reg', 'add',
                          'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                          '/V', 'DisableNotificationCenter', '/D', '0',
                          '/T', 'REG_DWORD', '/f'],
                         check=True, capture_output=True, text=True, encoding='cp866')
                except subprocess.CalledProcessError:
                    _run(['reg', 'delete',
                          'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                          '/v', 'DisableNotificationCenter', '/f'],
                         check=True, capture_output=True, text=True, encoding='cp866')
            else:
                _run(['reg', 'add',
                      'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                      '/V', 'DisableNotificationCenter', '/D', '1',
                      '/T', 'REG_DWORD', '/f'],
                     check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_taskbar_size(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                           '/v', 'TaskbarSmallIcons'],
                          capture_output=True, text=True, encoding='cp866')
            new_value = '0' if "0x1" in result.stdout else '1'
            _run(['reg', 'add',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                  '/V', 'TaskbarSmallIcons', '/D', new_value,
                  '/T', 'REG_DWORD', '/f'],
                 check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_search(self):
        try:
            current = self.check_search_state()
            new_value = {2: 1, 1: 0, 0: 2}[current]
            _run(['reg', 'add',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
                  '/V', 'SearchboxTaskbarMode', '/D', str(new_value),
                  '/T', 'REG_DWORD', '/f'],
                 check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_people(self):
        try:
            try:
                result = _run(['reg', 'query',
                               'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                               '/v', 'HidePeopleBar'],
                              capture_output=True, text=True, encoding='cp866')
                current_value = 1 if "0x1" in result.stdout else 0
            except Exception:
                current_value = 0

            if current_value == 1:
                try:
                    _run(['reg', 'delete',
                          'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                          '/v', 'HidePeopleBar', '/f'],
                         check=True, capture_output=True, text=True, encoding='cp866')
                except subprocess.CalledProcessError:
                    _run(['reg', 'add',
                          'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                          '/v', 'HidePeopleBar', '/d', '0', '/t', 'REG_DWORD', '/f'],
                         check=True, capture_output=True, text=True, encoding='cp866')
            else:
                _run(['reg', 'add',
                      'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                      '/v', 'HidePeopleBar', '/d', '1', '/t', 'REG_DWORD', '/f'],
                     check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_task_view(self):
        try:
            result = _run(['reg', 'query',
                           'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                           '/v', 'ShowTaskViewButton'],
                          capture_output=True, text=True, encoding='cp866')
            new_value = '1' if "0x0" in result.stdout else '0'
            _run(['reg', 'add',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                  '/V', 'ShowTaskViewButton', '/D', new_value,
                  '/T', 'REG_DWORD', '/f'],
                 check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_meet_now(self):
        try:
            try:
                result = _run(['reg', 'query',
                               'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                               '/v', 'HideSCAMeetNow'],
                              capture_output=True, text=True, encoding='cp866')
                current_value = 1 if "0x1" in result.stdout else 0
            except Exception:
                current_value = 0

            if current_value == 1:
                try:
                    _run(['reg', 'delete',
                          'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                          '/v', 'HideSCAMeetNow', '/f'],
                         check=True, capture_output=True, text=True, encoding='cp866')
                except subprocess.CalledProcessError:
                    _run(['reg', 'add',
                          'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                          '/v', 'HideSCAMeetNow', '/d', '0', '/t', 'REG_DWORD', '/f'],
                         check=True, capture_output=True, text=True, encoding='cp866')
            else:
                _run(['reg', 'add',
                      'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                      '/v', 'HideSCAMeetNow', '/d', '1', '/t', 'REG_DWORD', '/f'],
                     check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def toggle_grouping(self):
        try:
            current = self.check_grouping_state()
            new_value = {0: 1, 1: 2, 2: 0}[current]
            _run(['reg', 'add',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                  '/V', 'TaskbarGlomLevel', '/D', str(new_value),
                  '/T', 'REG_DWORD', '/f'],
                 check=True, capture_output=True, text=True, encoding='cp866')
        except Exception:
            pass

    def restart_explorer(self):
        try:
            _run(['taskkill', '/f', '/im', 'explorer.exe'],
                 capture_output=True, check=False)
            time.sleep(1)
            _popen(['explorer.exe'])
        except Exception:
            try:
                _popen(['explorer.exe'])
            except Exception:
                pass

    def reset_all_settings(self):
        try:
            settings_to_reset = [
                ('HKCU\\Control Panel\\International', 'sShortDate', 'dd.MM.yyyy', 'REG_SZ'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                 'TaskbarSmallIcons', '0', 'REG_DWORD'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
                 'SearchboxTaskbarMode', '2', 'REG_DWORD'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                 'ShowTaskViewButton', '1', 'REG_DWORD'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                 'TaskbarGlomLevel', '0', 'REG_DWORD'),
                ('HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                 'DisableNotificationCenter', '0', 'REG_DWORD'),
            ]

            policies_to_delete = [
                ('HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer', 'HidePeopleBar'),
            ]
            if self.is_windows_version_supported():
                policies_to_delete.append(
                    ('HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                     'HideSCAMeetNow'))

            for key, name, value, vtype in settings_to_reset:
                try:
                    _run(['reg', 'add', key, '/V', name, '/D', value, '/T', vtype, '/f'],
                         capture_output=True, check=True)
                except Exception:
                    continue

            for key, name in policies_to_delete:
                try:
                    _run(['reg', 'delete', key, '/v', name, '/f'],
                         capture_output=True, check=True)
                except Exception:
                    continue

            self.update_all_button_texts()
            self.restart_explorer()
        except Exception as e:
            messagebox.showerror("Ошибка сброса настроек",
                                 f"Не удалось сбросить настройки\nОшибка: {str(e)}")

    def update_all_button_texts(self):
        self.update_button_text(self.toggle_grouping_button,
                                "Изменить группировку кнопок",
                                self.check_grouping_state(), is_grouping=True)
        self.update_button_text(self.toggle_taskbar_size_button,
                                "Переключить размер значков панели задач",
                                self.check_taskbar_size_state())
        self.update_button_text(self.toggle_search_button, "Скрыть/Показать Поиск",
                                self.check_search_state(), is_search=True)
        self.update_button_text(self.toggle_task_view_button,
                                "Скрыть/Показать Просмотр задач",
                                not self.check_task_view_state())
        self.update_button_text(self.toggle_people_button, "Скрыть/Показать Люди",
                                not self.check_people_state())
        self.update_button_text(self.toggle_notification_center_button,
                                "Добавить/Убрать Центр уведомлений",
                                not self.check_notification_center_state())
        if self.is_windows_version_supported():
            self.update_button_text(self.toggle_meet_now_button,
                                    "Скрыть/Показать Провести собрание",
                                    not self.check_meet_now_state())
        self.update_button_text(self.toggle_weekday_button,
                                "Добавить/Убрать день недели в дате",
                                self.check_weekday_state())

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def hide(self):
        self.frame.pack_forget()