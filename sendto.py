import ctypes
import os
import tkinter as tk
from tkinter import ttk, messagebox

SENDTO_PATH = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'SendTo')


def _get_attrs(path):
    return ctypes.windll.kernel32.GetFileAttributesW(path)


def _set_attrs(path, attrs):
    return ctypes.windll.kernel32.SetFileAttributesW(path, attrs) != 0


def set_hidden(path, hidden=True):
    attrs = _get_attrs(path)
    if attrs == -1:
        return False
    attrs = (attrs | 0x2) if hidden else (attrs & ~0x2)
    return _set_attrs(path, attrs)


def is_hidden(path):
    attrs = _get_attrs(path)
    return attrs != -1 and bool(attrs & 0x2)


class SendToModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self.parent, text="Управление пунктами меню 'Отправить'",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        ctrl = ttk.Frame(self.parent)
        ctrl.pack(pady=5)
        buttons = ttk.Frame(ctrl)
        buttons.pack()

        self.toggle_all_btn = ttk.Button(buttons, text="", width=15,
                                         command=self.toggle_all)
        self.toggle_all_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.toggle_selected_btn = ttk.Button(buttons, text="Включить", width=15,
                                              command=self.toggle_selected)
        self.toggle_selected_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.listbox = tk.Listbox(self.parent, height=15, font=('Arial', 9))
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.listbox.bind('<<ListboxSelect>>', self._on_selection)

        self.refresh()

    def _items(self):
        if not os.path.exists(SENDTO_PATH):
            return []
        result = []
        for name in sorted(os.listdir(SENDTO_PATH)):
            if name.lower() == "desktop.ini":
                continue
            path = os.path.join(SENDTO_PATH, name)
            if os.path.isfile(path) or os.path.islink(path):
                result.append((name, path))
        return result

    def toggle_all(self):
        try:
            items = self._items()
            if not items:
                return
            all_hidden = all(is_hidden(p) for _, p in items)
            for _, path in items:
                set_hidden(path, not all_hidden)
            self.refresh()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пунктов: {e}")

    def _on_selection(self, event):
        sel = self.listbox.curselection()
        if sel:
            text = self.listbox.get(sel[0])
            self.toggle_selected_btn.config(
                text="Включить" if "❌" in text else "Отключить")

    def toggle_selected(self):
        try:
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showinfo("Информация", "Выберите пункт из списка")
                return
            text = self.listbox.get(sel[0])
            hidden = "❌" in text
            name = text[2:].replace(" (отключен)", "")
            path = os.path.join(SENDTO_PATH, name)
            if not os.path.exists(path):
                messagebox.showinfo("Информация", f"Пункт '{name}' не найден")
                return
            if not set_hidden(path, not hidden):
                messagebox.showerror("Ошибка",
                                     f"Не удалось изменить '{name}'")
                return
            self.refresh()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении: {e}")

    def refresh(self):
        self.listbox.delete(0, tk.END)
        if not os.path.exists(SENDTO_PATH):
            self.listbox.insert(0, "Папка SendTo не найдена")
            return
        items = self._items()
        if not items:
            self.listbox.insert(0, "Папка SendTo пуста")
            return

        visible_count = 0
        for name, path in items:
            if is_hidden(path):
                self.listbox.insert(tk.END, f"❌ {name} (отключен)")
            else:
                self.listbox.insert(tk.END, f"✅ {name}")
                visible_count += 1

        all_hidden = visible_count == 0
        self.toggle_all_btn.config(text="Включить всё" if all_hidden else "Отключить всё")