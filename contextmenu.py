"""Оболочка раздела «Контекстное меню».

Реэкспортирует публичный API пакета contextmenu_pkg, чтобы существующий код
(from contextmenu import check_menu_item_exists, CONTEXT_MENU_KEYS, ...)
продолжал работать без изменений.
"""
import tkinter as tk
from tkinter import ttk

from contextmenu_pkg import (
    CONTEXT_MENU_KEYS,
    ITEM_NAME_BY_KEY,
    check_menu_item_exists,
    DesktopMenuTab,
)
from contextmenu_pkg.createnew import CreateNewModule
from contextmenu_pkg.sendto import SendToModule

import logger

MOD = "contextmenu"


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.context_menu_module = ContextMenuModule(self.frame)
        self.show()

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def hide(self):
        self.frame.pack_forget()


class ContextMenuModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_new_module = CreateNewModule
        self.sendto_module = SendToModule
        self.create_widgets()
        self.show()

    def create_widgets(self):
        ttk.Label(
            self.parent,
            text="Настройка контекстного меню",
            font=('Arial', 16, 'bold'),
            foreground='#2c3e50'
        ).pack(pady=10)

        notebook = ttk.Notebook(self.parent)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        desktop_menu_frame = ttk.Frame(notebook)
        create_menu_frame = ttk.Frame(notebook)
        sendto_menu_frame = ttk.Frame(notebook)

        notebook.add(desktop_menu_frame, text="     Меню рабочего стола     ")
        notebook.add(create_menu_frame, text="          Меню Создать          ")
        notebook.add(sendto_menu_frame, text="         Меню Отправить         ")

        self.desktop_menu_tab = DesktopMenuTab(desktop_menu_frame)
        self.create_new_tab = self.create_new_module(create_menu_frame)
        self.sendto_tab = self.sendto_module(sendto_menu_frame)

    def show(self):
        self.parent.pack(fill=tk.BOTH, expand=True)

    def hide(self):
        self.parent.pack_forget()