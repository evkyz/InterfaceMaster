"""UI-хелперы: тултипы, сообщения с таймером, центрирование."""
import tkinter as tk
from tkinter import ttk


class Tooltip:
    """Универсальный тултип для любого виджета."""
    def __init__(self, widget, text, font=("Arial", 8), bg="lightyellow"):
        self.widget = widget
        self.text = text
        self.font = font
        self.bg = bg
        self._tip = None
        widget.bind("<Enter>", self._enter, add="+")
        widget.bind("<Leave>", self._leave, add="+")

    def _enter(self, event):
        if self._tip:
            return
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
            self._tip = tk.Toplevel(self.widget)
            self._tip.wm_overrideredirect(True)
            self._tip.wm_geometry(f"+{x}+{y}")
            tk.Label(self._tip, text=self.text, background=self.bg,
                     relief="solid", borderwidth=1, font=self.font,
                     padx=5, pady=2).pack()
        except Exception:
            self._tip = None

    def _leave(self, event):
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None

    def destroy(self):
        self._leave(None)


def add_tooltip(widget, text, **kwargs):
    """Создать тултип и привязать к виджету. Возвращает объект Tooltip."""
    return Tooltip(widget, text, **kwargs)


class MessageLabel:
    """Label с автопоказом сообщения на 3 секунды."""
    def __init__(self, parent, wraplength=400, font=('Arial', 9)):
        self.label = ttk.Label(parent, text="", font=font, wraplength=wraplength)
        self.timer = None
        self._parent = parent

    def pack(self, **kwargs):
        self.label.pack(**kwargs)

    def show(self, message, message_type="info", duration=3000):
        if self.timer:
            try:
                self._parent.after_cancel(self.timer)
            except Exception:
                pass
        colors = {"info": '#2c3e50', "success": '#27ae60', "warning": '#e74c3c'}
        try:
            if self.label.winfo_exists():
                self.label.config(text=message,
                                  foreground=colors.get(message_type, '#2c3e50'))
                self.label.pack(pady=5)
                self.timer = self._parent.after(duration, self.clear)
        except Exception:
            pass

    def clear(self):
        if self.timer:
            try:
                self._parent.after_cancel(self.timer)
            except Exception:
                pass
            self.timer = None
        try:
            if self.label.winfo_exists():
                self.label.config(text="")
                self.label.pack_forget()
        except Exception:
            pass


def center_window(root, width=None, height=None):
    root.update_idletasks()
    w = width or root.winfo_width()
    h = height or root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f'{w}x{h}+{x}+{y}')