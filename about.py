import hashlib
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes


class AboutModule:
    def __init__(self, parent, version, dll_info=None):
        self.parent = parent
        self.version = version

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

        # Эталонные хеши для модифицированных DLL (SHA-256)
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

    def get_dll_status_text(self):
        if not self.dll_info['dll_found']:
            return "Не найден"
        if self.dll_info['hash_match']:
            return "Совпадает"
        else:
            return "Не совпадает"

    def get_dll_status_color(self):
        if not self.dll_info['dll_found']:
            return '#e74c3c'
        if self.dll_info['hash_match']:
            return '#27ae60'
        else:
            return '#e67e22'

    def create_widgets(self):
        # Основной контейнер
        main_container = ttk.Frame(self.frame)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Заголовок
        title_label = ttk.Label(
            main_container,
            text="О программе",
            font=('Arial', 18, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=1)

        version_label = ttk.Label(
            main_container,
            text=f"Interface Master версия {self.version}",
            font=('Arial', 10),
            foreground='#3498db'
        )
        version_label.pack(pady=(0, 15))

        # Фрейм для информации
        info_frame = ttk.Frame(main_container)
        info_frame.pack(fill=tk.BOTH, expand=True)

        # Описание
        desc_frame = ttk.LabelFrame(info_frame, text="Описание", padding=10)
        desc_frame.pack(fill=tk.X, pady=(0, 10))

        desc_text = """Программа для настройки интерфейса Windows
Помощник в кастомизации внешнего вида системы"""

        desc_label = ttk.Label(
            desc_frame,
            text=desc_text,
            font=('Arial', 9),
            justify=tk.CENTER
        )
        desc_label.pack()

        # Возможности
        features_frame = ttk.LabelFrame(info_frame, text="Что можно настроить", padding=10)
        features_frame.pack(fill=tk.X, pady=(0, 15))

        features_text = """• Иконки рабочего стола          • Иконки дисков
• Стрелки ярлыков                    • Контекстное меню
• Панель задач                           • Проводник
                            • Модификации"""

        features_label = ttk.Label(
            features_frame,
            text=features_text,
            font=('Arial', 9),
            justify=tk.LEFT
        )
        features_label.pack()

        # Файл imaster.dll
        dll_frame = ttk.LabelFrame(info_frame, text="Файл imaster.dll", padding=10)
        dll_frame.pack(fill=tk.X, pady=(0, 15))

        self.create_dll_section(dll_frame)

        # Разработка
        dev_frame = ttk.LabelFrame(info_frame, text="Разработка", padding=10)
        dev_frame.pack(fill=tk.X)

        dev_text = f"""Идея: EvKyz
Реализация: DeepSeek AI
Сборка: 22 апреля 2026
Лицензия: MIT License"""

        dev_label = ttk.Label(
            dev_frame,
            text=dev_text,
            font=('Arial', 8),
            justify=tk.CENTER,
            foreground='#7f8c8d'
        )
        dev_label.pack()

        # Кнопка сброса внизу (без разделителя)
        reset_btn = ttk.Button(
            main_container,
            text="Сбросить все настройки",
            command=self.reset_all_settings,
            width=25
        )
        reset_btn.pack(side=tk.RIGHT, pady=(10, 5))
        
        # Подсказка
        self.create_tooltip(reset_btn, "Сбросить все настройки программы к стандартным значениям Windows")

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

    def create_dll_section(self, dll_frame):
        container = tk.Frame(dll_frame, bg='#f5f5f5', padx=5, pady=5)
        container.pack(fill=tk.BOTH, expand=True)

        status_color = self.get_dll_status_color()
        status_text = self.get_dll_status_text()

        status_label = tk.Label(
            container,
            text=f"Статус: {status_text}",
            font=('Consolas', 9, 'bold'),
            fg=status_color,
            bg='#f5f5f5',
            anchor='w'
        )
        status_label.pack(fill=tk.X, pady=(0, 5))

        if self.dll_info['dll_found'] and self.dll_info['dll_path']:
            path_label = tk.Label(
                container,
                text=f"Путь: {self.dll_info['dll_path']}",
                font=('Consolas', 8),
                bg='#f5f5f5',
                anchor='w'
            )
            path_label.pack(fill=tk.X, pady=(0, 3))

            if self.dll_info['actual_hash']:
                hash_label = tk.Label(
                    container,
                    text=f"Хеш SHA1: {self.dll_info['actual_hash']}",
                    font=('Consolas', 8),
                    bg='#f5f5f5',
                    anchor='w'
                )
                hash_label.pack(fill=tk.X)
        else:
            not_found_label = tk.Label(
                container,
                text="Файл imaster.dll не найден",
                font=('Consolas', 8),
                bg='#f5f5f5',
                anchor='w'
            )
            not_found_label.pack(fill=tk.X)

    def is_admin(self):
        """Проверка прав администратора"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False

    def log_to_file(self, message):
        """Записывает сообщение в файл reset.log на рабочем столе"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            log_file_path = os.path.join(desktop_path, "reset.log")
            os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
            with open(log_file_path, "a", encoding="utf-8") as log_file:
                log_file.write(message + "\n")
        except:
            pass

    def check_modified_files(self):
        """Проверяет, есть ли модифицированные файлы в системе"""
        modified_files = []
        for filename, expected_hash in self.mod_expected_hashes.items():
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            if os.path.exists(file_path):
                try:
                    with open(file_path, "rb") as f:
                        file_hash = hashlib.sha256(f.read()).hexdigest().upper()
                    if file_hash == expected_hash.upper():
                        modified_files.append(filename)
                except:
                    pass
        return modified_files

    def restore_modified_files(self, progress_callback=None):
        """Восстанавливает оригинальные системные DLL файлы"""
        restored = []
        
        for filename in self.mod_expected_hashes.keys():
            if progress_callback:
                progress_callback(f"Восстановление {filename}...")
            
            file_path = os.path.join(os.environ["WINDIR"], "System32", filename)
            backup_path = f"{file_path}_orig"
            
            # Проверяем, есть ли модифицированный файл
            is_modified = False
            if os.path.exists(file_path):
                try:
                    with open(file_path, "rb") as f:
                        current_hash = hashlib.sha256(f.read()).hexdigest().upper()
                    if current_hash == self.mod_expected_hashes[filename].upper():
                        is_modified = True
                except:
                    pass
            
            if is_modified:
                try:
                    # Завершаем explorer для разблокировки файлов
                    subprocess.run("taskkill /f /im explorer.exe", shell=True, 
                                 capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    # Удаляем модифицированный файл
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        self.log_to_file(f"Удалён модифицированный файл {filename}")
                    
                    # Восстанавливаем из бэкапа если есть
                    if os.path.exists(backup_path):
                        os.rename(backup_path, file_path)
                        self.log_to_file(f"Восстановлен оригинал {filename}")
                        restored.append(filename)
                except Exception as e:
                    self.log_to_file(f"Ошибка при восстановлении {filename}: {e}")
        
        return restored

    def reset_all_settings(self):
        """Сброс всех настроек программы к стандартным"""
        
        # Проверяем права администратора
        is_admin = self.is_admin()
        
        # Собираем информацию о модифицированных файлах
        modified_files = self.check_modified_files()
        
        # Формируем сообщение с предупреждением
        warning_msg = "ВНИМАНИЕ! Это действие сбросит все настройки к стандартным значениям Windows.\n\n"
        warning_msg += "Будут сброшены:\n"
        warning_msg += "• Иконки рабочего стола\n"
        warning_msg += "• Иконки дисков\n"
        warning_msg += "• Стрелки ярлыков\n"
        warning_msg += "• Текст '- ярлык'\n"
        warning_msg += "• Пункты контекстного меню\n"
        warning_msg += "• Пункты меню 'Создать'\n"
        warning_msg += "• Пункты меню 'Отправить'\n"
        warning_msg += "• Настройки Проводника\n"
        warning_msg += "• Настройки Панели задач\n"
        
        if modified_files:
            warning_msg += f"\n⚠️ БУДУТ ВОССТАНОВЛЕНЫ МОДИФИЦИРОВАННЫЕ ФАЙЛЫ:\n"
            for f in modified_files:
                warning_msg += f"  • {f}\n"
        
        warning_msg += "\n\nПродолжить?"
        
        response = messagebox.askyesno("Сброс всех настроек", warning_msg, icon='warning')
        if not response:
            return
        
        # Показываем окно прогресса
        progress_window = tk.Toplevel(self.parent)
        progress_window.title("Сброс настроек")
        progress_window.geometry("450x200")
        progress_window.resizable(False, False)
        progress_window.transient(self.parent)
        progress_window.grab_set()
        
        # Устанавливаем иконку окна
        try:
            # Ищем icon.ico в разных местах
            icon_paths = [
                os.path.join(os.path.dirname(sys.executable), "icon.ico") if getattr(sys, 'frozen', False) else None,
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico"),
                os.path.join(os.getcwd(), "icon.ico"),
                os.path.join(os.path.expanduser("~"), "Desktop", "icon.ico")
            ]
            for icon_path in icon_paths:
                if icon_path and os.path.exists(icon_path):
                    progress_window.iconbitmap(icon_path)
                    break
        except:
            pass
        
        # Центрируем окно
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (450 // 2)
        y = (progress_window.winfo_screenheight() // 2) - (200 // 2)
        progress_window.geometry(f"450x200+{x}+{y}")
        
        ttk.Label(
            progress_window,
            text="Сброс настроек...",
            font=('Arial', 12, 'bold')
        ).pack(pady=15)
        
        progress_bar = ttk.Progressbar(
            progress_window,
            mode='indeterminate',
            length=350
        )
        progress_bar.pack(pady=10)
        progress_bar.start(10)
        
        status_label = ttk.Label(
            progress_window,
            text="Выполняется сброс...",
            font=('Arial', 9),
            foreground='#7f8c8d'
        )
        status_label.pack(pady=10)
        
        errors = []
        
        def update_status(text):
            status_label.config(text=text)
            progress_window.update()
        
        def finalize():
            progress_bar.stop()
            progress_window.destroy()
            if errors:
                error_text = "Сброс завершён с ошибками:\n\n" + "\n".join(errors[:5])
                if len(errors) > 5:
                    error_text += f"\n\n... и ещё {len(errors)-5} ошибок"
                messagebox.showwarning("Сброс завершён с ошибками", error_text)
            else:
                messagebox.showinfo("Успех", "Все настройки успешно сброшены к стандартным значениям Windows!")
        
        try:
            # 1. Восстановление модифицированных DLL файлов (если есть)
            if modified_files:
                update_status("Восстановление системных файлов...")
                restored = self.restore_modified_files(update_status)
                if restored:
                    self.log_to_file(f"Восстановлены файлы: {', '.join(restored)}")
            
            # 2. Сброс иконок рабочего стола
            update_status("Сброс иконок рабочего стола...")
            try:
                from desktop import Module as DesktopModule
                temp_frame = ttk.Frame(progress_window)
                desktop_temp = DesktopModule(temp_frame, module_name="desktop")
                desktop_temp.reset_icons()
                desktop_temp.hide()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Иконки рабочего стола: {str(e)}")
                self.log_to_file(f"Ошибка: {e}")
            
            # 3. Сброс иконок дисков
            update_status("Сброс иконок дисков...")
            try:
                subprocess.run(
                    ['reg', 'delete', 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DriveIcons', '/f'],
                    capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW
                )
            except Exception as e:
                errors.append(f"Иконки дисков: {str(e)}")
            
            # 4. Сброс стрелок ярлыков и текста
            update_status("Сброс настроек ярлыков...")
            try:
                from str import Module as StrModule
                temp_frame = ttk.Frame(progress_window)
                str_temp = StrModule(temp_frame)
                str_temp.arrow_choice.set(1)
                str_temp.shortcut_text_choice.set(1)
                str_temp.apply_changes()
                str_temp.hide()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Настройки ярлыков: {str(e)}")
            
            # 5. Сброс контекстного меню
            update_status("Сброс контекстного меню...")
            try:
                from contextmenu import ContextMenuModule
                temp_frame = ttk.Frame(progress_window)
                contextmenu_temp = ContextMenuModule(temp_frame)
                # Получаем доступ к DesktopMenuTab и вызываем set_default_settings
                if hasattr(contextmenu_temp, 'desktop_menu_tab'):
                    contextmenu_temp.desktop_menu_tab.set_default_settings()
                elif hasattr(contextmenu_temp, 'desktop_menu_tab') and hasattr(contextmenu_temp.desktop_menu_tab, 'set_default_settings'):
                    contextmenu_temp.desktop_menu_tab.set_default_settings()
                else:
                    # Альтернативный способ - напрямую вызываем set_default_settings если он есть
                    if hasattr(contextmenu_temp, 'set_default_settings'):
                        contextmenu_temp.set_default_settings()
                contextmenu_temp.hide()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Контекстное меню: {str(e)}")
                self.log_to_file(f"Ошибка контекстного меню: {e}")
            
            # 6. Сброс меню "Создать" (только существующие пункты)
            update_status("Сброс меню 'Создать'...")
            try:
                from createnew import CreateNewModule
                temp_frame = ttk.Frame(progress_window)
                createnew_temp = CreateNewModule(temp_frame)
                items = createnew_temp.get_new_menu_items()
                for item in items:
                    if not item["enabled"]:
                        if item["ext"] == '.folder':
                            createnew_temp.enable_folder_menu_item()
                        else:
                            # Включаем только если файл существует (для Office)
                            try:
                                createnew_temp.enable_new_menu_item(item)
                            except:
                                pass
                createnew_temp.refresh_new_menu_list()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Меню 'Создать': {str(e)}")
            
            # 7. Сброс меню "Отправить"
            update_status("Сброс меню 'Отправить'...")
            try:
                from sendto import SendToModule
                temp_frame = ttk.Frame(progress_window)
                sendto_temp = SendToModule(temp_frame)
                sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'SendTo')
                if os.path.exists(sendto_path):
                    for item in os.listdir(sendto_path):
                        if item.lower() == "desktop.ini":
                            continue
                        item_path = os.path.join(sendto_path, item)
                        if os.path.isfile(item_path) or os.path.islink(item_path):
                            if sendto_temp.is_file_hidden(item_path):
                                sendto_temp.set_file_hidden(item_path, False)
                sendto_temp.refresh_sendto_list()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Меню 'Отправить': {str(e)}")
            
            # 8. Сброс настроек Проводника
            update_status("Сброс настроек Проводника...")
            try:
                from explorer import ExplorerModule
                temp_frame = ttk.Frame(progress_window)
                explorer_temp = ExplorerModule(temp_frame)
                # Включаем все папки
                for guid in explorer_temp.folders.keys():
                    explorer_temp.checkbox_vars[guid].set(True)
                explorer_temp.apply_folder_changes()
                # Включаем BitLocker
                if not explorer_temp.bitlocker_enabled:
                    explorer_temp.toggle_bitlocker()
                # Включаем Photo Viewer
                if not explorer_temp.photo_viewer_enabled:
                    explorer_temp.toggle_photo_viewer()
                # Скрываем расширения
                if explorer_temp.file_ext_enabled:
                    explorer_temp.toggle_file_ext()
                # Включаем недавние папки
                if not explorer_temp.recent_folders_enabled:
                    explorer_temp.toggle_recent_folders()
                explorer_temp.hide()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Проводник: {str(e)}")
            
            # 9. Сброс настроек Панели задач
            update_status("Сброс настроек Панели задач...")
            try:
                from taskbar import Module as TaskbarModule
                temp_frame = ttk.Frame(progress_window)
                taskbar_temp = TaskbarModule(temp_frame)
                taskbar_temp.reset_all_settings()
                taskbar_temp.hide()
                temp_frame.destroy()
            except Exception as e:
                errors.append(f"Панель задач: {str(e)}")
            
            # 10. Перезапуск проводника
            update_status("Перезапуск Проводника...")
            try:
                subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'], 
                             capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.Popen(['explorer.exe'], creationflags=subprocess.CREATE_NO_WINDOW)
            except:
                pass
            
            # Завершаем
            progress_window.after(500, finalize)
            
        except Exception as e:
            progress_bar.stop()
            progress_window.destroy()
            messagebox.showerror("Ошибка", f"Ошибка при сбросе настроек:\n{str(e)}")

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.frame.pack_forget()