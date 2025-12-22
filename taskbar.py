import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import sys


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.windows_version = self.get_windows_version()
        self.create_widgets()
        self.show()

    def get_windows_version(self):
        """Получает версию Windows"""
        try:
            # Используем более надежный способ получения версии Windows
            result = subprocess.run(
                ['wmic', 'os', 'get', 'Version', '/value'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            # Ищем версию в выводе
            import re
            match = re.search(r'Version=(\d+\.\d+\.\d+)', result.stdout)
            if match:
                return match.group(1)

            # Альтернативный способ через reg
            result = subprocess.run(
                ['reg', 'query', 'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion', '/v', 'CurrentBuildNumber'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            match = re.search(r'CurrentBuildNumber\s+REG_SZ\s+(\d+)', result.stdout)
            if match:
                return f"10.0.{match.group(1)}"

            return "10.0.0"
        except:
            return "10.0.0"

    def is_windows_version_supported(self):
        """Проверяет, поддерживается ли функция Провести собрание (Windows 21H2 и выше)"""
        try:
            # Получаем номер сборки из версии
            version_parts = self.windows_version.split('.')
            if len(version_parts) >= 3:
                build_number = int(version_parts[2])
                # Windows 21H2 имеет номер сборки 19044 и выше
                return build_number >= 19044
            return False
        except:
            return False

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

    def check_weekday_state(self):
        """Проверяет состояние отображения дня недели"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Control Panel\\International', '/v', 'sShortDate'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            return "ddd" in result.stdout
        except:
            return False

    def check_notification_center_state(self):
        """Проверяет состояние Центра уведомлений"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer', '/v',
                 'DisableNotificationCenter'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            return "0x1" in result.stdout
        except:
            return False

    def check_taskbar_size_state(self):
        """Проверяет состояние размера значков (большие/маленькие)"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/v',
                 'TaskbarSmallIcons'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            return "0x1" in result.stdout
        except:
            return False

    def check_search_state(self):
        """Проверяет состояние Поиска"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search', '/v',
                 'SearchboxTaskbarMode'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            # 0=скрыт, 1=значок, 2=поле
            if "0x0" in result.stdout:
                return 0  # Скрыто
            elif "0x1" in result.stdout:
                return 1  # Значок
            elif "0x2" in result.stdout:
                return 2  # Поле
            else:
                return 2  # По умолчанию - поле
        except:
            return 2  # По умолчанию - поле

    def check_people_state(self):
        """Проверяет состояние кнопки 'Люди'"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer', '/v', 'HidePeopleBar'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            return "0x1" in result.stdout
        except:
            return False

    def check_task_view_state(self):
        """Проверяет состояние кнопки 'Просмотр задач'"""
        try:
            result = subprocess.run(
                ['reg', 'query',
                 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/v',
                 'ShowTaskViewButton'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            return "0x0" in result.stdout
        except:
            return False

    def check_meet_now_state(self):
        """Проверяет состояние кнопки 'Провести собрание'"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer', '/v',
                 'HideSCAMeetNow'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            return "0x1" in result.stdout
        except:
            return False

    def check_grouping_state(self):
        """Проверяет состояние группировки кнопок"""
        try:
            result = subprocess.run(
                ['reg', 'query', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/v',
                 'TaskbarGlomLevel'],
                capture_output=True,
                text=True,
                encoding='cp866'
            )
            # 0=Всегда, 1=При заполнении, 2=Никогда
            if "0x0" in result.stdout:
                return 0  # Всегда
            elif "0x1" in result.stdout:
                return 1  # При заполнении
            elif "0x2" in result.stdout:
                return 2  # Никогда
            else:
                return 1  # По умолчанию - При заполнении
        except:
            return 1  # По умолчанию - При заполнении

    def update_button_text(self, button, base_text, state, is_search=False, is_grouping=False):
        """Обновляет текст кнопки с учетом состояния"""
        if button == self.toggle_taskbar_size_button:
            # Для кнопки изменения размера значков особая логика
            if state:
                button.config(text=f"{base_text} (Маленькие)")
            else:
                button.config(text=f"{base_text} (Большие)")
        elif is_search:
            # Для кнопки поиска специальная логика с тремя состояниями
            if state == 0:
                button.config(text=f"{base_text} (Скрыто)")
            elif state == 1:
                button.config(text=f"{base_text} (Значок)")
            else:  # state == 2
                button.config(text=f"{base_text} (Поле)")
        elif is_grouping:
            # Для кнопки группировки специальная логика с тремя состояниями
            if state == 0:
                button.config(text=f"{base_text} (Всегда)")
            elif state == 1:
                button.config(text=f"{base_text} (При заполнении)")
            else:  # state == 2
                button.config(text=f"{base_text} (Никогда)")
        else:
            # Для всех остальных кнопок
            if state:
                button.config(text=f"{base_text} (Показано)")
            else:
                button.config(text=f"{base_text} (Скрыто)")

    def create_widgets(self):
        """Создание виджетов для настройки панели задач"""
        title_label = ttk.Label(
            self.frame,
            text="Настройка Панели задач",
            font=('Arial', 14, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=5)

        desc_label = ttk.Label(
            self.frame,
            text="Настройте внешний вид и функциональность панели задач Windows",
            font=('Arial', 9),
            foreground='#7f8c8d'
        )
        desc_label.pack(pady=(0, 10))

        settings_frame = ttk.LabelFrame(self.frame, text="Доступные настройки", padding=10)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        # 1. Кнопка для группировки кнопок (3 состояния) - ПЕРВАЯ
        grouping_state = self.check_grouping_state()
        grouping_texts = {0: " (Всегда)", 1: " (При заполнении)", 2: " (Никогда)"}
        self.toggle_grouping_button = ttk.Button(
            settings_frame,
            text="Изменить группировку кнопок" + grouping_texts[grouping_state],
            command=lambda: self.toggle_grouping_with_update(),
            width=40
        )
        self.toggle_grouping_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_grouping_button,
                            "Циклически переключает: Всегда → При заполнении → Никогда → Всегда...")

        # 2. Кнопка для размера значков (особый случай) - ВТОРАЯ
        taskbar_size_state = self.check_taskbar_size_state()
        self.toggle_taskbar_size_button = ttk.Button(
            settings_frame,
            text="Переключить размер значков панели задач" + (" (Маленькие)" if taskbar_size_state else " (Большие)"),
            command=lambda: self.toggle_taskbar_size_with_update(),
            width=40
        )
        self.toggle_taskbar_size_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_taskbar_size_button,
                            "Переключает между большими и маленькими значками на панели задач")

        # 3. Кнопка для Поиска (3 состояния) - ТРЕТЬЯ
        search_state = self.check_search_state()
        search_texts = {0: " (Скрыто)", 1: " (Значок)", 2: " (Поле)"}
        self.toggle_search_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Поиск" + search_texts[search_state],
            command=lambda: self.toggle_search_with_update(),
            width=40
        )
        self.toggle_search_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_search_button, "Циклически переключает: Поле → Значок → Скрыто → Поле...")

        # 4. Кнопка для Просмотр задач - ЧЕТВЕРТАЯ
        task_view_state = self.check_task_view_state()
        self.toggle_task_view_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Просмотр задач" + (" (Показано)" if not task_view_state else " (Скрыто)"),
            command=lambda: self.toggle_task_view_with_update(),
            width=40
        )
        self.toggle_task_view_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_task_view_button,
                            "Скрывает или показывает кнопку 'Просмотр задач' на панели задач")

        # 5. Кнопка для Люди - ПЯТАЯ
        people_state = self.check_people_state()
        self.toggle_people_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Люди" + (" (Показано)" if not people_state else " (Скрыто)"),
            command=lambda: self.toggle_people_with_update(),
            width=40
        )
        self.toggle_people_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_people_button, "Скрывает или показывает кнопку 'Люди' на панели задач")

        # 6. Кнопка для Центра уведомлений - ШЕСТАЯ
        notification_state = self.check_notification_center_state()
        self.toggle_notification_center_button = ttk.Button(
            settings_frame,
            text="Добавить/Убрать Центр уведомлений" + (" (Показано)" if not notification_state else " (Скрыто)"),
            command=lambda: self.toggle_notification_center_with_update(),
            width=40
        )
        self.toggle_notification_center_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_notification_center_button, "Включает или отключает Центр уведомлений Windows")

        # 7. Кнопка для Провести собрание - СЕДЬМАЯ
        meet_now_state = self.check_meet_now_state() if self.is_windows_version_supported() else False
        self.toggle_meet_now_button = ttk.Button(
            settings_frame,
            text="Скрыть/Показать Провести собрание" + (" (Показано)" if not meet_now_state else " (Скрыто)"),
            command=lambda: self.toggle_meet_now_with_update(),
            width=40
        )
        self.toggle_meet_now_button.pack(pady=5, fill=tk.X)

        # Проверяем версию Windows и активируем/деактивируем кнопку
        if self.is_windows_version_supported():
            self.create_tooltip(self.toggle_meet_now_button,
                                "Скрывает или показывает кнопку 'Провести собрание' на панели задач")
        else:
            self.toggle_meet_now_button.config(state='disabled')
            self.create_tooltip(self.toggle_meet_now_button,
                                "Эта функция доступна только в Windows 21H2 (19044) и выше")

        # 8. Кнопка для дня недели - ПОСЛЕДНЯЯ
        weekday_state = self.check_weekday_state()
        self.toggle_weekday_button = ttk.Button(
            settings_frame,
            text="Добавить/Убрать день недели в дате" + (" (Показано)" if weekday_state else " (Скрыто)"),
            command=lambda: self.toggle_weekday_with_update(),
            width=40
        )
        self.toggle_weekday_button.pack(pady=5, fill=tk.X)
        self.create_tooltip(self.toggle_weekday_button, "Добавляет или убирает отображение дня недели в системной дате")

        extra_frame = ttk.Frame(settings_frame)
        extra_frame.pack(fill=tk.X, pady=10)

        restart_explorer_btn = ttk.Button(
            extra_frame,
            text="Перезапустить Проводник",
            command=self.restart_explorer,
            width=18
        )
        restart_explorer_btn.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        self.create_tooltip(restart_explorer_btn, "Перезагружает Проводник Windows для применения изменений")

        reset_all_btn = ttk.Button(
            extra_frame,
            text="Сбросить все настройки",
            command=self.reset_all_settings,
            width=18
        )
        reset_all_btn.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        self.create_tooltip(reset_all_btn, "Сбрасывает все настройки панели задач к значениям по умолчанию")

    # Оберточные функции для обновления текста кнопок после выполнения
    def toggle_weekday_with_update(self):
        self.toggle_weekday()
        new_state = self.check_weekday_state()
        self.update_button_text(self.toggle_weekday_button, "Добавить/Убрать день недели в дате", new_state)

    def toggle_notification_center_with_update(self):
        self.toggle_notification_center()
        new_state = self.check_notification_center_state()
        self.update_button_text(self.toggle_notification_center_button, "Добавить/Убрать Центр уведомлений",
                                not new_state)

    def toggle_taskbar_size_with_update(self):
        self.toggle_taskbar_size()
        new_state = self.check_taskbar_size_state()
        self.update_button_text(self.toggle_taskbar_size_button, "Переключить размер значков панели задач", new_state)

    def toggle_search_with_update(self):
        self.toggle_search()
        new_state = self.check_search_state()
        self.update_button_text(self.toggle_search_button, "Скрыть/Показать Поиск", new_state, is_search=True)

    def toggle_people_with_update(self):
        self.toggle_people()
        new_state = self.check_people_state()
        self.update_button_text(self.toggle_people_button, "Скрыть/Показать Люди", not new_state)

    def toggle_task_view_with_update(self):
        self.toggle_task_view()
        new_state = self.check_task_view_state()
        self.update_button_text(self.toggle_task_view_button, "Скрыть/Показать Просмотр задач", not new_state)

    def toggle_meet_now_with_update(self):
        self.toggle_meet_now()
        new_state = self.check_meet_now_state()
        self.update_button_text(self.toggle_meet_now_button, "Скрыть/Показать Провести собрание", not new_state)

    def toggle_grouping_with_update(self):
        self.toggle_grouping()
        new_state = self.check_grouping_state()
        self.update_button_text(self.toggle_grouping_button, "Изменить группировку кнопок", new_state, is_grouping=True)

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
        """Циклически переключить состояние Поиска: Поле → Значок → Скрыто → Поле..."""
        try:
            current_state = self.check_search_state()

            # Циклическое переключение: 2→1→0→2...
            if current_state == 2:  # Поле → Значок
                new_value = 1
            elif current_state == 1:  # Значок → Скрыто
                new_value = 0
            else:  # Скрыто → Поле
                new_value = 2

            subprocess.run(
                ['reg', 'add', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search', '/V',
                 'SearchboxTaskbarMode', '/D', str(new_value), '/T', 'REG_DWORD', '/f'],
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

    def toggle_meet_now(self):
        """Скрыть или показать кнопку 'Провести собрание' на панели задач."""
        try:
            # Проверяем текущее состояние
            try:
                result = subprocess.run(
                    ['reg', 'query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer', '/v',
                     'HideSCAMeetNow'],
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

            # Переключаем состояние
            if current_value == 1:
                # Если параметр существует и равен 1, удаляем его
                try:
                    subprocess.run(
                        ['reg', 'delete', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                         '/v', 'HideSCAMeetNow', '/f'],
                        check=True,
                        capture_output=True,
                        text=True,
                        encoding='cp866'
                    )
                except subprocess.CalledProcessError:
                    # Если не удалось удалить, устанавливаем значение 0
                    subprocess.run(
                        ['reg', 'add', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                         '/v', 'HideSCAMeetNow', '/d', '0', '/t', 'REG_DWORD', '/f'],
                        check=True,
                        capture_output=True,
                        text=True,
                        encoding='cp866'
                    )
            else:
                # Если параметр не существует или равен 0, устанавливаем 1
                subprocess.run(
                    ['reg', 'add', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                     '/v', 'HideSCAMeetNow', '/d', '1', '/t', 'REG_DWORD', '/f'],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='cp866'
                )

        except subprocess.CalledProcessError:
            pass  # Не показываем сообщение об ошибке
        except Exception:
            pass  # Не показываем сообщение об ошибке

    def toggle_grouping(self):
        """Циклически переключить состояние группировки кнопок: Всегда → При заполнении → Никогда → Всегда..."""
        try:
            current_state = self.check_grouping_state()

            # Циклическое переключение: 0→1→2→0...
            if current_state == 0:  # Всегда → При заполнении
                new_value = 1
            elif current_state == 1:  # При заполнении → Никогда
                new_value = 2
            else:  # Никогда → Всегда
                new_value = 0

            subprocess.run(
                ['reg', 'add', 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', '/V',
                 'TaskbarGlomLevel', '/D', str(new_value), '/T', 'REG_DWORD', '/f'],
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
                ['HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced', 'TaskbarGlomLevel', '0',
                 'REG_DWORD'],
            ]

            notification_center_settings = [
                ['HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer', 'DisableNotificationCenter', '0',
                 'REG_DWORD'],
            ]

            policies_to_delete = [
                ['HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer', 'HidePeopleBar'],
            ]

            # Добавляем настройку для MeetNow в сброс, если она поддерживается
            if self.is_windows_version_supported():
                policies_to_delete.append(
                    ['HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer', 'HideSCAMeetNow']
                )

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

            # Обновляем тексты кнопок после сброса
            self.update_all_button_texts()

            # Показываем всплывающее окно об успешном сбросе
            messagebox.showinfo(
                "Сброс настроек",
                "Настройки сброшены по умолчанию"
            )

        except Exception as e:
            # Показываем сообщение об ошибке
            messagebox.showerror(
                "Ошибка сброса настроек",
                f"Не удалось сбросить настройки.\nОшибка: {str(e)}"
            )

    def update_all_button_texts(self):
        """Обновляет тексты всех кнопок после сброса настроек"""
        # Группировка кнопок
        grouping_state = self.check_grouping_state()
        self.update_button_text(self.toggle_grouping_button, "Изменить группировку кнопок", grouping_state,
                                is_grouping=True)

        # Размер значков
        taskbar_size_state = self.check_taskbar_size_state()
        self.update_button_text(self.toggle_taskbar_size_button, "Переключить размер значков панели задач",
                                taskbar_size_state)

        # Поиск
        search_state = self.check_search_state()
        self.update_button_text(self.toggle_search_button, "Скрыть/Показать Поиск", search_state, is_search=True)

        # Просмотр задач
        task_view_state = self.check_task_view_state()
        self.update_button_text(self.toggle_task_view_button, "Скрыть/Показать Просмотр задач", not task_view_state)

        # Люди
        people_state = self.check_people_state()
        self.update_button_text(self.toggle_people_button, "Скрыть/Показать Люди", not people_state)

        # Центр уведомлений
        notification_state = self.check_notification_center_state()
        self.update_button_text(self.toggle_notification_center_button, "Добавить/Убрать Центр уведомлений",
                                not notification_state)

        # Провести собрание
        if self.is_windows_version_supported():
            meet_now_state = self.check_meet_now_state()
            self.update_button_text(self.toggle_meet_now_button, "Скрыть/Показать Провести собрание",
                                    not meet_now_state)

        # День недели
        weekday_state = self.check_weekday_state()
        self.update_button_text(self.toggle_weekday_button, "Добавить/Убрать день недели в дате", weekday_state)

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()