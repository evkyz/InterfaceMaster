import subprocess
import tkinter as tk
from tkinter import ttk, messagebox


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.create_widgets()
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
        """Создание виджетов для настройки панели задач"""
        title_label = ttk.Label(
            self.frame,
            text="Настройка Панели задач",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        desc_label = ttk.Label(
            self.frame,
            text="Настройте внешний вид и функциональность панели задач Windows",
            font=('Arial', 10),
            foreground='#7f8c8d'
        )
        desc_label.pack(pady=(0, 20))

        settings_frame = ttk.LabelFrame(self.frame, text="Доступные настройки", padding=15)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.toggle_weekday_button = ttk.Button(
            settings_frame,
            text="Добавить/Убрать день недели в дате",
            command=self.toggle_weekday,
            width=35
        )
        self.toggle_weekday_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.toggle_weekday_button, "Добавляет или убирает отображение дня недели в системной дате")

        self.toggle_notification_center_button = ttk.Button(
            settings_frame,
            text="Добавить/Убрать Центр уведомлений",
            command=self.toggle_notification_center,
            width=35
        )
        self.toggle_notification_center_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.toggle_notification_center_button, "Включает или отключает Центр уведомлений Windows")

        self.toggle_taskbar_size_button = ttk.Button(
            settings_frame,
            text="Переключить размер значков панели задач",
            command=self.toggle_taskbar_size,
            width=35
        )
        self.toggle_taskbar_size_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.toggle_taskbar_size_button,
                            "Переключает между большими и маленькими значками на панели задач")

        self.toggle_search_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Поиск",
            command=self.toggle_search,
            width=35
        )
        self.toggle_search_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.toggle_search_button, "Скрывает или показывает поле поиска на панели задач")

        self.toggle_people_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Люди",
            command=self.toggle_people,
            width=35
        )
        self.toggle_people_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.toggle_people_button, "Скрывает или показывает кнопку 'Люди' на панели задач")

        self.toggle_task_view_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Просмотр задач",
            command=self.toggle_task_view,
            width=35
        )
        self.toggle_task_view_button.pack(pady=8, fill=tk.X)
        self.create_tooltip(self.toggle_task_view_button,
                            "Скрывает или показывает кнопку 'Просмотр задач' на панели задач")

        extra_frame = ttk.Frame(settings_frame)
        extra_frame.pack(fill=tk.X, pady=15)

        restart_explorer_btn = ttk.Button(
            extra_frame,
            text="Перезапустить Проводник",
            command=self.restart_explorer,
            width=20
        )
        restart_explorer_btn.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        self.create_tooltip(restart_explorer_btn, "Немедленно перезагружает Проводник Windows для применения изменений")

        reset_all_btn = ttk.Button(
            extra_frame,
            text="Сбросить все настройки",
            command=self.reset_all_settings,
            width=20
        )
        reset_all_btn.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        self.create_tooltip(reset_all_btn, "Сбрасывает все настройки панели задач к значениям по умолчанию")

        info_label = ttk.Label(
            settings_frame,
            font=('Arial', 8),
            foreground='#e67e22',
            justify=tk.CENTER
        )
        info_label.pack(pady=10)

    def toggle_weekday(self):
        """Добавить или убрать день недели в формате даты."""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Control Panel\\International', '/v', 'sShortDate'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )

            if "ddd" in result.stdout:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Control Panel\\International', '/V', 'sShortDate',
                     '/D', 'dd.MM.yyyy', '/T', 'REG_SZ', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )
            else:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Control Panel\\International', '/V', 'sShortDate',
                     '/D', 'ddd dd.MM.yyyy', '/T', 'REG_SZ', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def toggle_notification_center(self):
        """Включить или отключить Центр уведомлений."""
        try:
            try:
                result = subprocess.run(
                    ['reg', 'query', 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer', '/v',
                     'DisableNotificationCenter'],
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

                if "0x1" in result.stdout:
                    current_value = 1
                else:
                    current_value = 0
            except subprocess.CalledProcessError:
                current_value = 0

            if current_value == 1:
                try:
                    subprocess.run(
                        ['reg', 'add', 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer', '/V',
                         'DisableNotificationCenter', '/D', '0', '/T', 'REG_DWORD', '/f'],
                        check=True,
                        capture_output=True,
                        text=True,
                        encoding='cp866'
                    )
                except subprocess.CalledProcessError:
                    subprocess.run(
                        ['reg', 'delete', 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                         '/v', 'DisableNotificationCenter', '/f'],
                        check=True,
                        capture_output=True,
                        text=True,
                        encoding='cp866'
                    )
            else:
                subprocess.run(
                    ['reg', 'add', 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer', '/V',
                     'DisableNotificationCenter', '/D', '1', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def toggle_taskbar_size(self):
        """Переключить размер значков на панели задач (большие/маленькие)."""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/v',
                 'TaskbarSmallIcons'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )

            if "0x1" in result.stdout:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/V',
                     'TaskbarSmallIcons', '/D', '0', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )
            else:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/V',
                     'TaskbarSmallIcons', '/D', '1', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def toggle_search(self):
        """Скрыть или показать Поиск на панели задач."""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search', '/v',
                 'SearchboxTaskbarMode'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )

            if "0x2" in result.stdout:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search', '/V',
                     'SearchboxTaskbarMode', '/D', '0', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )
            else:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search', '/V',
                     'SearchboxTaskbarMode', '/D', '2', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def toggle_people(self):
        """Скрыть или показать Люди на панели задач."""
        try:
            try:
                result = subprocess.run(
                    ['reg', 'query', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer', '/v', 'HidePeopleBar'],
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )
                if "0x1" in result.stdout:
                    current_value = 1
                else:
                    current_value = 0
            except subprocess.CalledProcessError:
                current_value = 0

            if current_value == 1:
                try:
                    subprocess.run(
                        ['reg', 'delete', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                         '/v', 'HidePeopleBar', '/f'],
                        check=True,
                        capture_output=True,
                        text=True,
                        encoding='cp866'
                    )
                except subprocess.CalledProcessError:
                    subprocess.run(
                        ['reg', 'add', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                         '/v', 'HidePeopleBar', '/d', '0', '/t', 'REG_DWORD', '/f'],
                        check=True,
                        capture_output=True,
                        text=True,
                        encoding='cp866'
                    )
            else:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                     '/v', 'HidePeopleBar', '/d', '1', '/t', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def toggle_task_view(self):
        """Скрыть или показать Просмотр задач на панели задач."""
        try:
            result = subprocess.run(
                ['reg', 'query',
                 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/v',
                 'ShowTaskViewButton'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )

            if "0x0" in result.stdout:
                subprocess.run(
                    ['reg', 'add',
                     'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/V',
                     'ShowTaskViewButton', '/D', '1', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )
            else:
                subprocess.run(
                    ['reg', 'add',
                     'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/V',
                     'ShowTaskViewButton', '/D', '0', '/T', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def restart_explorer(self):
        """Перезапустить Проводник Windows."""
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                           capture_output=True, check=True)

            subprocess.Popen(['explorer.exe'])

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def reset_all_settings(self):
        """Сбросить все настройки панели задач к значениям по умолчанию."""
        try:
            settings_to_reset = [
                ['HKCU\\Control Panel\\International', 'sShortDate', 'dd.MM.yyyy', 'REG_SZ'],
                ['HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', 'TaskbarSmallIcons', '0',
                 'REG_DWORD'],
                ['HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search', 'SearchboxTaskbarMode', '2',
                 'REG_DWORD'],
                ['HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', 'ShowTaskViewButton', '1',
                 'REG_DWORD'],
            ]

            notification_center_settings = [
                ['HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer', 'DisableNotificationCenter', '0',
                 'REG_DWORD'],
            ]

            policies_to_delete = [
                ['HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer', 'HidePeopleBar'],
            ]

            for setting in settings_to_reset:
                try:
                    subprocess.run(
                        ['reg', 'add', setting[0], '/V', setting[1], '/D', setting[2], '/T', setting[3], '/f'],
                        capture_output=True,
                        check=True
                    )
                except subprocess.CalledProcessError:
                    continue

            for setting in notification_center_settings:
                try:
                    subprocess.run(
                        ['reg', 'add', setting[0], '/V', setting[1], '/D', setting[2], '/T', setting[3], '/f'],
                        capture_output=True,
                        check=True
                    )
                except subprocess.CalledProcessError:
                    try:
                        subprocess.run(
                            ['reg', 'delete', setting[0], '/v', setting[1], '/f'],
                            capture_output=True,
                            check=True
                        )
                    except subprocess.CalledProcessError:
                        continue

            for policy in policies_to_delete:
                try:
                    subprocess.run(
                        ['reg', 'delete', policy[0], '/v', policy[1], '/f'],
                        capture_output=True,
                        check=True
                    )
                except subprocess.CalledProcessError:
                    continue

        except Exception:
            pass  # Не показываем сообщение об ошибке

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()