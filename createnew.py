# createnew.py
import os
import subprocess
import tkinter as tk
import winreg
from tkinter import ttk, messagebox


class CreateNewModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        main_container = ttk.Frame(self.parent)
        main_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_container)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>",
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        ttk.Label(scrollable_frame, text="Управление пунктами меню 'Создать'",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        control_frame = ttk.Frame(scrollable_frame)
        control_frame.pack(pady=5)
        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)
        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        self.new_menu_toggle_all_btn = ttk.Button(
            buttons_frame, text="", command=self.toggle_all_new_menu_items, width=15)
        self.new_menu_toggle_all_btn.pack(side=tk.LEFT, padx=5, pady=5)

        ttk.Separator(scrollable_frame, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)

        self.new_menu_items_frame = ttk.Frame(scrollable_frame)
        self.new_menu_items_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.refresh_new_menu_list()

    def toggle_all_new_menu_items(self):
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
        try:
            for widget in self.new_menu_items_frame.winfo_children():
                widget.destroy()

            items = self.get_new_menu_items()
            if not items:
                ttk.Label(self.new_menu_items_frame, text="Пункты меню 'Создать' не найдены",
                          font=('Arial', 10), foreground='#7f8c8d').pack(pady=20)
                return

            all_disabled = all(not item["enabled"] for item in items)
            self.new_menu_toggle_all_btn.config(text="Включить всё" if all_disabled else "Отключить всё")

            header_frame = ttk.Frame(self.new_menu_items_frame)
            header_frame.pack(fill=tk.X, pady=(0, 5))
            ttk.Label(header_frame, text="Название пункта", font=('Arial', 9, 'bold'), width=30).pack(
                side=tk.LEFT, padx=5)
            ttk.Label(header_frame, text="Статус", font=('Arial', 9, 'bold'), width=15).pack(
                side=tk.LEFT, padx=5)
            ttk.Label(header_frame, text="Действие", font=('Arial', 9, 'bold'), width=15).pack(
                side=tk.LEFT, padx=0)

            for item in items:
                self.create_new_menu_item_row(item)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обновлении списка: {str(e)}")

    def create_new_menu_item_row(self, item):
        try:
            row_frame = ttk.Frame(self.new_menu_items_frame, relief='groove', borderwidth=1)
            row_frame.pack(fill=tk.X, pady=2, padx=5)

            ttk.Label(row_frame, text=item["name"], width=29, anchor='w').pack(
                side=tk.LEFT, padx=5, pady=5)

            status_text = "✅ Включен" if item["enabled"] else "❌ Отключен"
            ttk.Label(row_frame, text=status_text, width=15,
                      foreground='#27ae60' if item["enabled"] else '#e74c3c').pack(
                side=tk.LEFT, padx=5, pady=5)

            action_text = "Отключить" if item["enabled"] else "Включить"
            ttk.Button(row_frame, text=action_text, width=15,
                       command=lambda i=item: self.toggle_new_menu_item(
                           i["ext"], i.get("path"), i.get("hive"), i.get("base_path"))).pack(
                side=tk.LEFT, padx=5, pady=5)
        except Exception:
            pass

    def get_new_menu_items(self):
        items = []
        try:
            # Папка (HKCR)
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Folder\ShellNew'):
                    items.append({"ext": '.folder', "name": "Папка", "enabled": True,
                                  "path": r'Folder\ShellNew',
                                  "hive": winreg.HKEY_CLASSES_ROOT,
                                  "base_path": r'Folder\ShellNew'})
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Folder\ShellNew_no'):
                        items.append({"ext": '.folder', "name": "Папка", "enabled": False,
                                      "path": r'Folder\ShellNew_no',
                                      "hive": winreg.HKEY_CLASSES_ROOT,
                                      "base_path": r'Folder\ShellNew'})
                except FileNotFoundError:
                    items.append({"ext": '.folder', "name": "Папка", "enabled": False,
                                  "path": None, "hive": winreg.HKEY_CLASSES_ROOT,
                                  "base_path": r'Folder\ShellNew'})

            target_extensions = {
                '.bmp': r'SOFTWARE\Classes\.bmp\ShellNew',
                '.contact': r'SOFTWARE\Classes\.contact\ShellNew',
                '.lnk': r'SOFTWARE\Classes\.lnk\ShellNew',
                '.rtf': r'SOFTWARE\Classes\.rtf\ShellNew',
                '.txt': r'SOFTWARE\Classes\.txt\ShellNew',
                '.zip': r'SOFTWARE\Classes\.zip\CompressedFolder\ShellNew'
            }

            for ext, base_path in target_extensions.items():
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path):
                        items.append({"ext": ext, "name": self.get_file_type_name(ext),
                                      "enabled": True, "path": base_path,
                                      "hive": winreg.HKEY_LOCAL_MACHINE,
                                      "base_path": base_path})
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path + '_no'):
                            items.append({"ext": ext, "name": self.get_file_type_name(ext),
                                          "enabled": False, "path": base_path + '_no',
                                          "hive": winreg.HKEY_LOCAL_MACHINE,
                                          "base_path": base_path})
                    except FileNotFoundError:
                        items.append({"ext": ext, "name": self.get_file_type_name(ext),
                                      "enabled": False, "path": None,
                                      "hive": winreg.HKEY_LOCAL_MACHINE,
                                      "base_path": base_path})

            items.extend(self.get_office_menu_items())
        except Exception as e:
            print(f"Ошибка: {e}")
        return items

    def get_office_menu_items(self):
        items = []
        try:
            office_version = None
            for version_path, version in [
                (r'SOFTWARE\Microsoft\Office\16.0', '16.0'),
                (r'SOFTWARE\Microsoft\Office\15.0', '15.0'),
                (r'SOFTWARE\Microsoft\Office\14.0', '14.0'),
                (r'SOFTWARE\Microsoft\Office\12.0', '12.0'),
                (r'SOFTWARE\Microsoft\Office\11.0', '11.0')
            ]:
                for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                    try:
                        with winreg.OpenKey(hive, version_path):
                            office_version = version
                            break
                    except Exception:
                        continue
                if office_version:
                    break

            if not office_version:
                return items

            short_version = office_version.replace('.0', '')

            office_defs = [
                ('.docx', "Документ Word", "Word.Document", winreg.HKEY_LOCAL_MACHINE),
                ('.xlsx', "Книга Excel", "Excel.Sheet", winreg.HKEY_LOCAL_MACHINE),
                ('.pptx', "Презентация PowerPoint", "PowerPoint.Show", winreg.HKEY_LOCAL_MACHINE),
                ('.pub', "Публикация Publisher", "Publisher.Document", winreg.HKEY_CLASSES_ROOT),
            ]

            for ext, name, class_name, hive in office_defs:
                base_path = f"SOFTWARE\\Classes\\{ext}\\{class_name}.{short_version}\\ShellNew"
                found = False
                for candidate in [
                    f"{ext}\\{class_name}.{short_version}\\ShellNew",
                    f"{ext}\\{class_name}.{office_version}\\ShellNew",
                    f"SOFTWARE\\Classes\\{ext}\\{class_name}.{short_version}\\ShellNew",
                    f"SOFTWARE\\Classes\\{ext}\\{class_name}.{office_version}\\ShellNew",
                ]:
                    if not candidate.startswith("SOFTWARE"):
                        key_hive = winreg.HKEY_CLASSES_ROOT
                    else:
                        key_hive = winreg.HKEY_LOCAL_MACHINE
                    try:
                        with winreg.OpenKey(key_hive, candidate):
                            items.append({"ext": ext, "name": name, "enabled": True,
                                          "path": candidate, "hive": key_hive,
                                          "base_path": candidate})
                            found = True
                            break
                    except Exception:
                        continue

                if not found:
                    items.append({"ext": ext, "name": name, "enabled": False,
                                  "path": None, "hive": hive,
                                  "base_path": base_path})
        except Exception as e:
            print(f"Ошибка получения Office пунктов: {e}")
        return items

    def get_file_type_name(self, ext):
        return {
            '.folder': 'Папка', '.bmp': 'Рисунок BMP', '.contact': 'Контакт',
            '.lnk': 'Ярлык', '.rtf': 'Документ RTF', '.txt': 'Текстовый документ',
            '.zip': 'Архив ZIP', '.docx': 'Документ Word', '.xlsx': 'Книга Excel',
            '.pptx': 'Презентация PowerPoint', '.pub': 'Публикация Publisher'
        }.get(ext, f"Файл {ext.lstrip('.')}")

    def toggle_new_menu_item(self, ext, path=None, hive=winreg.HKEY_LOCAL_MACHINE, base_path=None):
        try:
            if ext == '.folder':
                base_path = base_path or r'Folder\ShellNew'
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, base_path):
                        self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, base_path, base_path + '_no')
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, base_path + '_no'):
                            self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, base_path + '_no', base_path)
                    except FileNotFoundError:
                        try:
                            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, base_path)
                            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, base_path, 0,
                                                winreg.KEY_SET_VALUE) as key:
                                winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                        except Exception:
                            pass
            else:
                base_path = base_path or f'SOFTWARE\\Classes\\{ext}\\ShellNew'
                try:
                    with winreg.OpenKey(hive, base_path):
                        self.rename_registry_key(hive, base_path, base_path + '_no')
                except FileNotFoundError:
                    try:
                        with winreg.OpenKey(hive, base_path + '_no'):
                            self.rename_registry_key(hive, base_path + '_no', base_path)
                    except FileNotFoundError:
                        try:
                            parent_path = '\\'.join(base_path.split('\\')[:-1])
                            winreg.CreateKey(hive, parent_path)
                            winreg.CreateKey(hive, base_path)
                            with winreg.OpenKey(hive, base_path, 0, winreg.KEY_SET_VALUE) as key:
                                winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
                        except Exception:
                            pass
            self.refresh_new_menu_list()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении: {str(e)}")

    def enable_folder_menu_item(self):
        try:
            folder_path = r'Folder\ShellNew'
            try:
                self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, folder_path + '_no', folder_path)
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path):
                        pass
                except FileNotFoundError:
                    winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, folder_path)
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path, 0,
                                        winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
        except Exception as e:
            print(f"Ошибка включения папки: {e}")

    def disable_folder_menu_item(self):
        try:
            folder_path = r'Folder\ShellNew'
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, folder_path):
                self.rename_registry_key(winreg.HKEY_CLASSES_ROOT, folder_path, folder_path + '_no')
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"Ошибка отключения папки: {e}")

    def enable_new_menu_item(self, item):
        try:
            hive = item.get("hive", winreg.HKEY_LOCAL_MACHINE)
            base_path = item.get("base_path")
            if not base_path:
                return
            try:
                with winreg.OpenKey(hive, base_path + '_no'):
                    self.rename_registry_key(hive, base_path + '_no', base_path)
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(hive, base_path):
                        pass
                except FileNotFoundError:
                    parent_path = '\\'.join(base_path.split('\\')[:-1])
                    winreg.CreateKey(hive, parent_path)
                    winreg.CreateKey(hive, base_path)
                    with winreg.OpenKey(hive, base_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.SetValueEx(key, "NullFile", 0, winreg.REG_SZ, "")
        except Exception as e:
            print(f"Ошибка включения {item.get('ext')}: {e}")

    def disable_new_menu_item(self, item):
        try:
            hive = item.get("hive", winreg.HKEY_LOCAL_MACHINE)
            base_path = item.get("base_path")
            if not base_path:
                return
            with winreg.OpenKey(hive, base_path):
                self.rename_registry_key(hive, base_path, base_path + '_no')
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"Ошибка отключения {item.get('ext')}: {e}")

    def rename_registry_key(self, hive, old_key_path, new_key_path):
        self.copy_registry_key(hive, old_key_path, new_key_path)
        self.delete_registry_key(hive, old_key_path)

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