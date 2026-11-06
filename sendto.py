import os
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes


class SendToModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def set_file_hidden(self, filepath, hidden=True):
        try:
            attrs = ctypes.windll.kernel32.GetFileAttributesW(filepath)
            if attrs == -1:
                return False
            if hidden:
                attrs |= 0x2
            else:
                attrs &= ~0x2
            return ctypes.windll.kernel32.SetFileAttributesW(filepath, attrs) != 0
        except Exception:
            return False

    def is_file_hidden(self, filepath):
        try:
            attrs = ctypes.windll.kernel32.GetFileAttributesW(filepath)
            if attrs == -1:
                return False
            return bool(attrs & 0x2)
        except Exception:
            return False

    def create_widgets(self):
        ttk.Label(self.parent, text="Управление пунктами меню 'Отправить'",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        control_frame = ttk.Frame(self.parent)
        control_frame.pack(pady=5)
        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)
        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        self.sendto_toggle_all_btn = ttk.Button(
            buttons_frame, text="", command=self.toggle_all_sendto_items, width=15)
        self.sendto_toggle_all_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.toggle_selected_btn = ttk.Button(
            buttons_frame, text="Включить",
            command=self.toggle_selected_sendto_item, width=15)
        self.toggle_selected_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.sendto_listbox = tk.Listbox(self.parent, height=15, font=('Arial', 9))
        self.sendto_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.sendto_listbox.bind('<<ListboxSelect>>', self.on_sendto_selection)

        self.refresh_sendto_list()

    def toggle_all_sendto_items(self):
        try:
            sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft',
                                       'Windows', 'SendTo')
            if not os.path.exists(sendto_path):
                messagebox.showerror("Ошибка", "Папка SendTo не найдена")
                return

            items = os.listdir(sendto_path)
            has_hidden = False
            has_visible = False
            for item in items:
                if item.lower() == "desktop.ini":
                    continue
                item_path = os.path.join(sendto_path, item)
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if self.is_file_hidden(item_path):
                        has_hidden = True
                    else:
                        has_visible = True

            all_hidden = has_hidden and not has_visible

            for item in items:
                if item.lower() == "desktop.ini":
                    continue
                item_path = os.path.join(sendto_path, item)
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if all_hidden:
                        if self.is_file_hidden(item_path):
                            self.set_file_hidden(item_path, False)
                    else:
                        if not self.is_file_hidden(item_path):
                            self.set_file_hidden(item_path, True)

            self.refresh_sendto_list()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пунктов: {str(e)}")

    def on_sendto_selection(self, event):
        selection = self.sendto_listbox.curselection()
        if selection:
            item_text = self.sendto_listbox.get(selection[0])
            self.toggle_selected_btn.config(
                text="Включить" if "❌" in item_text else "Отключить")

    def toggle_selected_sendto_item(self):
        try:
            selection = self.sendto_listbox.curselection()
            if not selection:
                messagebox.showinfo("Информация", "Выберите пункт из списка")
                return

            item_text = self.sendto_listbox.get(selection[0])
            sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft',
                                       'Windows', 'SendTo')

            if "❌" in item_text:
                original_name = item_text[2:].replace(" (отключен)", "")
                item_path = os.path.join(sendto_path, original_name)
                if not os.path.exists(item_path):
                    messagebox.showinfo("Информация",
                                        f"Пункт '{original_name}' не найден")
                    return
                if not self.is_file_hidden(item_path):
                    messagebox.showinfo("Информация",
                                        f"Пункт '{original_name}' уже видим")
                    return
                if not self.set_file_hidden(item_path, False):
                    messagebox.showerror("Ошибка",
                                         f"Не удалось сделать видимым пункт '{original_name}'")
                    return
                self.refresh_sendto_list()
            else:
                original_name = item_text[2:] if "✅" in item_text else item_text
                item_path = os.path.join(sendto_path, original_name)
                if not os.path.exists(item_path):
                    messagebox.showinfo("Информация",
                                        f"Пункт '{original_name}' не найден")
                    return
                if self.is_file_hidden(item_path):
                    messagebox.showinfo("Информация",
                                        f"Пункт '{original_name}' уже скрыт")
                    return
                if not self.set_file_hidden(item_path, True):
                    messagebox.showerror("Ошибка",
                                         f"Не удалось скрыть пункт '{original_name}'")
                    return
                self.refresh_sendto_list()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пункта: {str(e)}")

    def refresh_sendto_list(self):
        try:
            self.sendto_listbox.delete(0, tk.END)
            sendto_path = os.path.join(os.environ['APPDATA'], 'Microsoft',
                                       'Windows', 'SendTo')
            if not os.path.exists(sendto_path):
                self.sendto_listbox.insert(0, "Папка SendTo не найдена")
                return

            items = os.listdir(sendto_path)
            if not items:
                self.sendto_listbox.insert(0, "Папка SendTo пуста")
                return

            has_hidden = False
            has_visible = False
            display_items = []

            for item in sorted(items):
                if item.lower() == "desktop.ini":
                    continue
                item_path = os.path.join(sendto_path, item)
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    if self.is_file_hidden(item_path):
                        display_items.append(f"❌ {item} (отключен)")
                        has_hidden = True
                    else:
                        display_items.append(f"✅ {item}")
                        has_visible = True

            if has_visible and not has_hidden:
                self.sendto_toggle_all_btn.config(text="Отключить всё")
            elif has_hidden and not has_visible:
                self.sendto_toggle_all_btn.config(text="Включить всё")
            else:
                self.sendto_toggle_all_btn.config(text="Отключить всё")

            for text in display_items:
                self.sendto_listbox.insert(tk.END, text)
        except Exception as e:
            self.sendto_listbox.insert(0, f"Ошибка: {e}")