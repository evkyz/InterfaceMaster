import os
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes


class SendToModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def set_file_hidden(self, filepath, hidden=True):
        """Установка/снятие атрибута 'Скрытый' для файла"""
        try:
            # Получаем текущие атрибуты файла
            attrs = ctypes.windll.kernel32.GetFileAttributesW(filepath)
            if attrs == -1:  # Ошибка получения атрибутов
                return False
            
            if hidden:
                # Устанавливаем скрытый атрибут
                attrs |= 0x2  # FILE_ATTRIBUTE_HIDDEN
            else:
                # Снимаем скрытый атрибут
                attrs &= ~0x2
            
            # Применяем новые атрибуты
            result = ctypes.windll.kernel32.SetFileAttributesW(filepath, attrs)
            return result != 0
        except Exception:
            return False

    def is_file_hidden(self, filepath):
        """Проверка, скрыт ли файл"""
        try:
            attrs = ctypes.windll.kernel32.GetFileAttributesW(filepath)
            if attrs == -1:
                return False
            return bool(attrs & 0x2)  # FILE_ATTRIBUTE_HIDDEN
        except Exception:
            return False

    def create_widgets(self):
        """Создание вкладки меню Отправить"""

        title_label = ttk.Label(
            self.parent,
            text="Управление пунктами меню 'Отправить'",
            font=('Arial', 12, 'bold'),
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)

        control_frame = ttk.Frame(self.parent)
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

        self.sendto_listbox = tk.Listbox(self.parent, height=15, font=('Arial', 9))
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

            has_hidden_items = False
            has_visible_items = False

            for item in items:
                if item.lower() == "desktop.ini":
                    continue

                item_path = os.path.join(sendto_path, item)

                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if self.is_file_hidden(item_path):
                        has_hidden_items = True
                    else:
                        has_visible_items = True

            # Если все видимые элементы скрыты, то включаем всё, иначе выключаем всё
            all_hidden = has_hidden_items and not has_visible_items

            for item in items:
                item_path = os.path.join(sendto_path, item)

                if item.lower() == "desktop.ini":
                    continue

                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if all_hidden:
                        # Включаем скрытые элементы (делаем видимыми)
                        if self.is_file_hidden(item_path):
                            self.set_file_hidden(item_path, False)
                    else:
                        # Выключаем видимые элементы (делаем скрытыми)
                        if not self.is_file_hidden(item_path):
                            self.set_file_hidden(item_path, True)

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
                # Элемент скрыт, нужно сделать видимым
                original_name = item_text[2:].replace(" (отключен)", "")
                item_path = os.path.join(sendto_path, original_name)

                try:
                    if os.path.exists(item_path):
                        if self.is_file_hidden(item_path):
                            if self.set_file_hidden(item_path, False):
                                self.refresh_sendto_list()
                            else:
                                messagebox.showerror("Ошибка", f"Не удалось сделать видимым пункт '{original_name}'")
                        else:
                            messagebox.showinfo("Информация", f"Пункт '{original_name}' уже видим")
                    else:
                        messagebox.showinfo("Информация", f"Пункт '{original_name}' не найден")
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось включить пункт: {str(e)}")
            else:
                # Элемент видим, нужно сделать скрытым
                if "✅" in item_text:
                    original_name = item_text[2:]
                else:
                    original_name = item_text

                item_path = os.path.join(sendto_path, original_name)

                try:
                    if os.path.exists(item_path):
                        if not self.is_file_hidden(item_path):
                            if self.set_file_hidden(item_path, True):
                                self.refresh_sendto_list()
                            else:
                                messagebox.showerror("Ошибка", f"Не удалось скрыть пункт '{original_name}'")
                        else:
                            messagebox.showinfo("Информация", f"Пункт '{original_name}' уже скрыт")
                    else:
                        messagebox.showinfo("Информация", f"Пункт '{original_name}' не найден")
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

            has_hidden_items = False
            has_visible_items = False

            display_items = []

            for item in sorted(items):
                if item.lower() == "desktop.ini":
                    continue

                item_path = os.path.join(sendto_path, item)

                if os.path.isfile(item_path) or os.path.islink(item_path):
                    is_hidden = self.is_file_hidden(item_path)
                    
                    if is_hidden:
                        display_text = f"❌ {item} (отключен)"
                        has_hidden_items = True
                    else:
                        display_text = f"✅ {item}"
                        has_visible_items = True

                    display_items.append((item, display_text, is_hidden))

            # Обновляем текст кнопки "Включить всё/Отключить всё"
            if has_visible_items and not has_hidden_items:
                self.sendto_toggle_all_btn.config(text="Отключить всё")
            elif has_hidden_items and not has_visible_items:
                self.sendto_toggle_all_btn.config(text="Включить всё")
            elif has_hidden_items and has_visible_items:
                # Если есть и скрытые, и видимые, по умолчанию предлагаем отключить видимые
                self.sendto_toggle_all_btn.config(text="Отключить всё")
            else:
                self.sendto_toggle_all_btn.config(text="Отключить всё")

            for _, display_text, _ in display_items:
                self.sendto_listbox.insert(tk.END, display_text)

        except Exception as e:
            self.sendto_listbox.insert(0, f"Ошибка: {e}")