import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import platform


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)

        self.content_frame = ttk.Frame(self.frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self.arrow_choice = tk.IntVar(value=0)
        self.windows_version = self._get_windows_version()

        self.create_widgets()
        self.show()

    def _get_windows_version(self):
        """Получить версию Windows"""
        try:
            return platform.version()
        except:
            return ""

    def _is_windows_10_19045_or_higher(self):
        """Проверить, является ли версия Windows 10.0.19045 или выше"""
        try:
            version_parts = self.windows_version.split('.')
            if len(version_parts) >= 3:
                major = int(version_parts[0])
                minor = int(version_parts[1])
                build = int(version_parts[2])

                if major == 10 and minor == 0:
                    return build >= 19045
            return False
        except:
            return False

    def create_widgets(self):
        """Создание виджетов для настройки стрелок ярлыков"""
        title_label = ttk.Label(
            self.content_frame,
            text="Настройка стрелок ярлыков",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        desc_label = ttk.Label(
            self.content_frame,
            text="Выберите тип отображения стрелок на ярлыках Windows",
            font=('Arial', 10),
            foreground='#7f8c8d'
        )
        desc_label.pack(pady=(0, 20))

        settings_frame = ttk.LabelFrame(self.content_frame, text="Выбор типа стрелок", padding=15)
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.big_arrow_check = ttk.Radiobutton(
            settings_frame,
            text="Большие стрелки (стандартные)",
            variable=self.arrow_choice,
            value=1
        )
        self.big_arrow_check.pack(anchor=tk.W, pady=5)

        self.small_arrow_check = ttk.Radiobutton(
            settings_frame,
            text="Маленькие стрелки",
            variable=self.arrow_choice,
            value=2
        )
        self.small_arrow_check.pack(anchor=tk.W, pady=5)

        self.no_arrow_check = ttk.Radiobutton(
            settings_frame,
            text="Убрать стрелки",
            variable=self.arrow_choice,
            value=3
        )
        self.no_arrow_check.pack(anchor=tk.W, pady=5)

        button_frame = ttk.Frame(settings_frame)
        button_frame.pack(fill=tk.X, pady=15)

        apply_btn = ttk.Button(
            button_frame,
            text="Применить настройки",
            command=self.apply_changes,
            width=25
        )
        apply_btn.pack(pady=(0, 10))

        restart_btn = ttk.Button(
            button_frame,
            text="Перезапустить Проводник",
            command=self.restart_explorer,
            width=25
        )
        restart_btn.pack(pady=5)

        self.status_label = ttk.Label(
            settings_frame,
            text="",
            font=('Arial', 9),
            foreground='#27ae60',
            wraplength=400,
            justify=tk.CENTER
        )
        self.status_label.pack(pady=5)

    def _delete_registry_key_if_exists(self):
        """Удалить ключ реестра, если он существует."""
        try:
            check_result = subprocess.run(
                ['reg', 'query', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons', '/v',
                 '29'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if check_result.returncode == 0:
                subprocess.run(
                    ['reg', 'delete', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
                     '/v', '29', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    check=True
                )
        except:
            pass

    def _set_registry_value(self, value):
        """Установить значение в реестре."""
        try:
            subprocess.run(
                ['reg', 'add', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Icons',
                 '/v', '29', '/d', value, '/t', 'REG_SZ', '/f'],
                capture_output=True,
                text=True,
                encoding='cp866',
                creationflags=subprocess.CREATE_NO_WINDOW,
                check=True
            )
        except subprocess.CalledProcessError as e:
            raise e

    def _apply_windows10_fix(self):
        """Применить фикс для Windows 10/11, чтобы убрать стрелки без белых квадратов"""
        try:
            # Метод для новой версии Windows
            self._set_registry_value('')

            # Дополнительный параметр для новых версий Windows
            try:
                subprocess.run(
                    ['reg', 'add', 'HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer',
                     '/v', 'Link', '/t', 'REG_BINARY', '/d', '00000000', '/f'],
                    capture_output=True,
                    text=True,
                    encoding='cp866',
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    check=True
                )
            except:
                pass

            return True
        except:
            return False

    def _clear_icon_cache(self):
        """Очистить кэш иконок Windows"""
        try:
            # Только очистка кэша без перезапуска explorer
            subprocess.run(['ie4uinit.exe', '-show'],
                           capture_output=True,
                           text=True,
                           encoding='cp866',
                           creationflags=subprocess.CREATE_NO_WINDOW)
            return True
        except:
            return False

    def apply_changes(self):
        """Применить выбранные изменения."""
        choice = self.arrow_choice.get()

        try:
            if choice == 1:
                # Большие стрелки
                self._delete_registry_key_if_exists()
                # Удаляем дополнительный ключ
                try:
                    subprocess.run(
                        ['reg', 'delete', 'HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer',
                         '/v', 'Link', '/f'],
                        capture_output=True,
                        text=True,
                        encoding='cp866',
                        creationflags=subprocess.CREATE_NO_WINDOW,
                        check=True
                    )
                except:
                    pass

                self.show_status("✓ Установлены стандартные большие стрелки")

            elif choice == 2:
                # Маленькие стрелки
                self._set_registry_value('C:\\Windows\\System32\\shell32.dll,-30')
                self.show_status("✓ Установлены маленькие стрелки")

            elif choice == 3:
                # Убрать стрелки
                if self._is_windows_10_19045_or_higher():
                    # Для Windows 10 22H2 и выше
                    if self._apply_windows10_fix():
                        self.show_status("✓ Стрелки ярлыков удалены")
                    else:
                        # Если не сработало, пробуем старый метод
                        self._set_registry_value('C:\\Windows\\System32\\shell32.dll,-50')
                        self.show_status("✓ Стрелки ярлыков удалены")
                else:
                    # Для старых версий
                    self._set_registry_value('C:\\Windows\\System32\\shell32.dll,-50')
                    self.show_status("✓ Стрелки ярлыков удалены")

            # Очищаем кэш иконок после изменений
            self._clear_icon_cache()

        except subprocess.CalledProcessError:
            messagebox.showerror("Ошибка", "Не удалось изменить")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка: {str(e)}")

    def restart_explorer(self):
        """Перезагрузить Проводник Windows."""
        try:
            subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'],
                           capture_output=True,
                           text=True,
                           encoding='cp866',
                           creationflags=subprocess.CREATE_NO_WINDOW)

            subprocess.Popen('explorer.exe', shell=True,
                             creationflags=subprocess.CREATE_NO_WINDOW)

            self.show_status("✓ Проводник успешно перезапущен")

        except Exception:
            messagebox.showerror("Ошибка", "Не удалось перезапустить Проводник")

    def show_status(self, message, is_success=True):
        """Показать статус операции"""
        color = '#27ae60' if is_success else '#e74c3c'
        self.status_label.config(text=message, foreground=color)

        self.frame.after(5000, lambda: self.status_label.config(text=""))

    def show(self):
        """Показать модуль"""
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        """Скрыть модуль"""
        self.frame.pack_forget()