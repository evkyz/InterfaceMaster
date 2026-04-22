# createnew.py
import os
import subprocess
import tkinter as tk
import winreg
from tkinter import ttk, messagebox
import sys


class CreateNewModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        """Создание вкладки меню Создать"""

        main_container = ttk.Frame(self.parent)
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

            all_disabled = all(not item["enabled"] for item in items)

            for item in items:
                if all_disabled:
                    if not item["enabled"]:
                        if item["ext"] == '.folder':
                            self.enable_folder_menu_item()
                        else:
                            self.enable_new_menu_item(item)
                else:
                    if item["enabled"]:
                        if item["ext"] == '.folder':
                            self.disable_folder_menu_item()
                        else:
                            self.disable_new_menu_item(item)

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
                command=lambda i=item: self.toggle_new_menu_item(i["ext"], i.get("path"), i.get("hive"), i.get("base_path"))
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
        """Получение списка пунктов меню Создать"""
        items = []
        try:
            # Папка (HKCR)
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Folder\ShellNew'):
                    items.append({
                        "ext": '.folder',
                        "name": "Папка",
                        "enabled": True,
                        "path": r'Folder\ShellNew',
                        "hive": winreg.HKEY_CLASSES_ROOT,
                        "base_path": r'Folder\ShellNew'
                    })
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Folder\ShellNew_no'):
                        items.append({
                            "ext": '.folder',
                            "name": "Папка",
                            "enabled": False,
                            "path": r'Folder\ShellNew_no',
                            "hive": winreg.HKEY_CLASSES_ROOT,
                            "base_path": r'Folder\ShellNew'
                        })
                except FileNotFoundError:
                    items.append({
                        "ext": '.folder',
                        "name": "Папка",
                        "enabled": False,
                        "path": None,
                        "hive": winreg.HKEY_CLASSES_ROOT,
                        "base_path": r'Folder\ShellNew'
                    })

            # Стандартные расширения (HKLM)
            target_extensions = {
                '.bmp': r'SOFTWARE\Classes\.bmp\ShellNew',
                '.contact': r'SOFTWARE\Classes\.contact\ShellNew',
                '.lnk': r'SOFTWARE\Classes\.lnk\ShellNew',
                '.rtf': r'SOFTWARE\Classes\.rtf\ShellNew',
                '.txt': r'SOFTWARE\Classes\.txt\ShellNew',
                '.zip': r'SOFTWARE\Classes\.zip\CompressedFolder\ShellNew'
            }

            for ext, base_path in target_extensions.items():
                shellnew_path = base_path
                shellnew_no_path = base_path + '_no'

                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_path):
                        items.append({
                            "ext": ext,
                            "name": self.get_file_type_name(ext),
                            "enabled": True,
                            "path": shellnew_path,
                            "hive": winreg.HKEY_LOCAL_MACHINE,
                            "base_path": base_path
                        })
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, shellnew_no_path):
                            items.append({
                                "ext": ext,
                                "name": self.get_file_type_name(ext),
                                "enabled": False,
                                "path": shellnew_no_path,
                                "hive": winreg.HKEY_LOCAL_MACHINE,
                                "base_path": base_path
                            })
                    except FileNotFoundError:
                        items.append({
                            "ext": ext,
                            "name": self.get_file_type_name(ext),
                            "enabled": False,
                            "path": None,
                            "hive": winreg.HKEY_LOCAL_MACHINE,
                            "base_path": base_path
                        })

            # Office расширения
            office_items = self.get_office_menu_items()
            items.extend(office_items)

        except Exception as e:
            print(f"Ошибка: {e}")

        return items

    def get_office_menu_items(self):
        """Получение пунктов меню для Microsoft Office"""
        items = []
        
        try:
            # Определяем версию Office
            office_version = None
            office_versions = [
                (r'SOFTWARE\Microsoft\Office\16.0', '16.0'),
                (r'SOFTWARE\Microsoft\Office\15.0', '15.0'),
                (r'SOFTWARE\Microsoft\Office\14.0', '14.0'),
                (r'SOFTWARE\Microsoft\Office\12.0', '12.0'),
                (r'SOFTWARE\Microsoft\Office\11.0', '11.0')
            ]

            for version_path, version in office_versions:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, version_path):
                        office_version = version
                        break
                except:
                    continue

            if not office_version:
                for version_path, version in office_versions:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, version_path):
                            office_version = version
                            break
                    except:
                        continue

            if office_version:
                # Word
                word_paths = [
                    f"SOFTWARE\\Classes\\.docx\\Word.Document.{office_version}\\ShellNew",
                    f"SOFTWARE\\Classes\\.docx\\Word.Document.{office_version.replace('.0', '')}\\ShellNew",
                    r"SOFTWARE\Classes\.docx\Word.Document.12\ShellNew"
                ]
                word_found = False
                for path in word_paths:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                            items.append({
                                "ext": '.docx',
                                "name": "Документ Word",
                                "enabled": True,
                                "path": path,
                                "hive": winreg.HKEY_LOCAL_MACHINE,
                                "base_path": path
                            })
                            word_found = True
                            break
                    except:
                        continue
                
                if not word_found:
                    items.append({
                        "ext": '.docx',
                        "name": "Документ Word",
                        "enabled": False,
                        "path": None,
                        "hive": winreg.HKEY_LOCAL_MACHINE,
                        "base_path": f"SOFTWARE\\Classes\\.docx\\Word.Document.{office_version.replace('.0', '')}\\ShellNew"
                    })

                # Excel
                excel_paths = [
                    f"SOFTWARE\\Classes\\.xlsx\\Excel.Sheet.{office_version}\\ShellNew",
                    f"SOFTWARE\\Classes\\.xlsx\\Excel.Sheet.{office_version.replace('.0', '')}\\ShellNew",
                    r"SOFTWARE\Classes\.xlsx\Excel.Sheet.12\ShellNew"
                ]
                excel_found = False
                for path in excel_paths:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                            items.append({
                                "ext": '.xlsx',
                                "name": "Книга Excel",
                                "enabled": True,
                                "path": path,
                                "hive": winreg.HKEY_LOCAL_MACHINE,
                                "base_path": path
                            })
                            excel_found = True
                            break
                    except:
                        continue
                
                if not excel_found:
                    items.append({
                        "ext": '.xlsx',
                        "name": "Книга Excel",
                        "enabled": False,
                        "path": None,
                        "hive": winreg.HKEY_LOCAL_MACHINE,
                        "base_path": f"SOFTWARE\\Classes\\.xlsx\\Excel.Sheet.{office_version.replace('.0', '')}\\ShellNew"
                    })

                # PowerPoint
                ppt_paths = [
                    f"SOFTWARE\\Classes\\.pptx\\PowerPoint.Show.{office_version}\\ShellNew",
                    f"SOFTWARE\\Classes\\.pptx\\PowerPoint.Show.{office_version.replace('.0', '')}\\ShellNew",
                    r"SOFTWARE\Classes\.pptx\PowerPoint.Show.12\ShellNew"
                ]
                ppt_found = False
                for path in ppt_paths:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                            items.append({
                                "ext": '.pptx',
                                "name": "Презентация PowerPoint",
                                "enabled": True,
                                "path": path,
                                "hive": winreg.HKEY_LOCAL_MACHINE,
                                "base_path": path
                            })
                            ppt_found = True
                            break
                    except:
                        continue
                
                if not ppt_found:
                    items.append({
                        "ext": '.pptx',
                        "name": "Презентация PowerPoint",
                        "enabled": False,
                        "path": None,
                        "hive": winreg.HKEY_LOCAL_MACHINE,
                        "base_path": f"SOFTWARE\\Classes\\.pptx\\PowerPoint.Show.{office_version.replace('.0', '')}\\ShellNew"
                    })

                # Publisher - используем версию БЕЗ .0
                pub_version = office_version.replace('.0', '')
                
                pub_paths = [
                    f".pub\\Publisher.Document.{pub_version}\\ShellNew",
                    f".pub\\Publisher.Document.{office_version}\\ShellNew"
                ]
                pub_found = False
                for path in pub_paths:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, path):
                            items.append({
                                "ext": '.pub',
                                "name": "Публикация Publisher",
                                "enabled": True,
                                "path": path,
                                "hive": winreg.HKEY_CLASSES_ROOT,
                                "base_path": path
                            })
                            pub_found = True
                            break
                    except:
                        continue
                
                if not pub_found:
                    # Проверяем через HKLM
                    for path in [f"SOFTWARE\\Classes\\.pub\\Publisher.Document.{pub_version}\\ShellNew",
                                 f"SOFTWARE\\Classes\\.pub\\Publisher.Document.{office_version}\\ShellNew"]:
                        try:
                            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
                                items.append({
                                    "ext": '.pub',
                                    "name": "Публикация Publisher",
                                    "enabled": True,
                                    "path": path,
                                    "hive": winreg.HKEY_LOCAL_MACHINE,
                                    "base_path": path
                                })
                                pub_found = True
                                break
                        except:
                            continue
                
                if not pub_found:
                    # Добавляем как отключенный с правильным базовым путём (без .0)
                    items.append({
                        "ext": '.pub',
                        "name": "Публикация Publisher",
                        "enabled": False,
                        "path": None,
                        "hive": winreg.HKEY_CLASSES_ROOT,
                        "base_path": f".pub\\Publisher.Document.{pub_version}\\ShellNew"
                    })

        except Exception as e:
            print(f"Ошибка получения Office пунктов: {e}")

        return items

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
        return basic_names.get(ext, f"Файл {ext.lstrip('.')}")

    def toggle_new_menu_item(self, ext, path=None, hive=winreg.HKEY_LOCAL_MACHINE, base_path=None):
        """Переключение состояния пункта меню Создать"""
        try:
            if ext == '.folder':
                # Папка
                if base_path is None:
                    base_path = r'Folder\ShellNew'
                
                old_key = base_path
                new_key = base_path + '_no'
                
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, old_key):
                        self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, old_key, new_key)
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, new_key):
                            self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, new_key, old_key)
                    except FileNotFoundError:
                        try:
                            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, old_key)
                            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, old_key, 0, winreg.KEY_SET_VALUE) as key:
                                winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                        except:
                            pass
            else:
                # Файлы
                if base_path is None:
                    base_path = f'SOFTWARE\\Classes\\{ext}\\ShellNew'
                
                old_key = base_path
                new_key = base_path + '_no'
                
                try:
                    with winreg.OpenKey(hive, old_key):
                        self.rename_registry_key(hive, old_key, new_key)
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(hive, new_key):
                            self.rename_registry_key(hive, new_key, old_key)
                    except FileNotFoundError:
                        try:
                            # Создаем родительские ключи если нужно
                            parent_path = '\\'.join(base_path.split('\\')[:-1])
                            try:
                                winreg.CreateKey(hive, parent_path)
                            except:
                                pass
                            winreg.CreateKey(hive, old_key)
                            with winreg.OpenKey(hive, old_key, 0, winreg.KEY_SET_VALUE) as key:
                                winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                        except:
                            pass

            self.refresh_new_menu_list()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении: {str(e)}")

    def enable_folder_menu_item(self):
        """Включение папки"""
        try:
            folder_path = r'Folder\ShellNew'
            folder_no_path = folder_path + '_no'

            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_no_path):
                    self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, folder_no_path, folder_path)
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path):
                        pass
                except FileNotFoundError:
                    winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, folder_path)
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")

        except Exception as e:
            print(f"Ошибка включения папки: {e}")

    def disable_folder_menu_item(self):
        """Отключение папки"""
        try:
            folder_path = r'Folder\ShellNew'
            folder_no_path = folder_path + '_no'

            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path):
                    self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, folder_path, folder_no_path)
            except FileNotFoundError:
                pass

        except Exception as e:
            print(f"Ошибка отключения папки: {e}")

    def enable_new_menu_item(self, item):
        """Включение пункта файла"""
        try:
            hive = item.get("hive", winreg.HKEY_LOCAL_MACHINE)
            base_path = item.get("base_path")
            
            if not base_path:
                return
            
            old_key = base_path
            new_key = base_path + '_no'
            
            try:
                with winreg.OpenKey(hive, new_key):
                    self.rename_registry_key(hive, new_key, old_key)
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(hive, old_key):
                        pass
                except FileNotFoundError:
                    # Создаем родительские ключи если нужно
                    parent_path = '\\'.join(base_path.split('\\')[:-1])
                    try:
                        winreg.CreateKey(hive, parent_path)
                    except:
                        pass
                    winreg.CreateKey(hive, old_key)
                    with winreg.OpenKey(hive, old_key, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")

        except Exception as e:
            print(f"Ошибка включения {item.get('ext')}: {e}")

    def disable_new_menu_item(self, item):
        """Отключение пункта файла"""
        try:
            hive = item.get("hive", winreg.HKEY_LOCAL_MACHINE)
            base_path = item.get("base_path")
            
            if not base_path:
                return
            
            old_key = base_path
            new_key = base_path + '_no'
            
            try:
                with winreg.OpenKey(hive, old_key):
                    self.rename_registry_key(hive, old_key, new_key)
            except FileNotFoundError:
                pass

        except Exception as e:
            print(f"Ошибка отключения {item.get('ext')}: {e}")

    def rename_registry_key(self, hive, old_key_path, new_key_path):
        """Переименование ключа реестра"""
        try:
            self.copy_registry_key(hive, old_key_path, new_key_path)
            self.delete_registry_key(hive, old_key_path)
        except Exception as e:
            raise e

    def copy_registry_key(self, hive, source_path, dest_path):
        """Копирование ключа реестра"""
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
        except:
            raise