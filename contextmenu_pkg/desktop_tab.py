"""UI-вкладка «Меню рабочего стола»."""
import tkinter as tk
from tkinter import ttk, messagebox

import logger
from core import win_utils
from core.ui_utils import add_tooltip
from contextmenu_pkg import menu_checks, menu_power, menu_shell_order
from contextmenu_pkg.desktop_items import build_menu_items

MOD = "contextmenu"


class DesktopMenuTab:
    TOOLTIP_KEEP_NAMES = {
        "Диспетчер задач",
        "Очистить буфер",
        "Перезапустить explorer",
        "Редактор реестра",
        "Выключение/Перезагрузка/Гибернация",
        "Администрирование",
        "Система",
    }

    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        canvas = tk.Canvas(self.parent)
        scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        ttk.Label(scrollable_frame, text="Управление пунктами контекстного меню",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        control_frame = ttk.Frame(scrollable_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        center_container = ttk.Frame(control_frame)
        center_container.pack(expand=True)
        buttons_frame = ttk.Frame(center_container)
        buttons_frame.pack()

        self.restart_explorer_btn = ttk.Button(
            buttons_frame, text="Перезапустить Проводник",
            command=self.restart_explorer, width=25)
        self.restart_explorer_btn.pack(side=tk.LEFT, padx=5, pady=5)
        self.create_tooltip(self.restart_explorer_btn,
                            "Перезапустить проводник Windows")

        default_btn = ttk.Button(
            buttons_frame, text="По умолчанию",
            command=self.set_default_settings, width=15)
        default_btn.pack(side=tk.LEFT, padx=5, pady=5)
        self.create_tooltip(default_btn,
                            "Сбросить пункты контекстного меню к стандартным")

        self.menu_items_data = build_menu_items()

        self.menu_item_buttons = []
        for item in self.menu_items_data:
            button_frame = ttk.Frame(scrollable_frame)
            button_frame.pack(fill=tk.X, padx=10, pady=5)

            btn = self.create_menu_item_button(button_frame, item)
            self.menu_item_buttons.append({"button": btn, "data": item})

            supported_check = item.get("supported_check")
            if supported_check is not None and not logger.is_debug():
                if not supported_check():
                    btn.config(state="disabled")
                    tip_text = item.get("unsupported_tip",
                                        "Недоступно на этой версии")
                    if menu_checks.is_ltsc_edition():
                        tip_text = "Недоступно в LTSC редакциях"
                    self.create_tooltip(btn, tip_text)
                    continue

            if item.get("add_tooltip") and item["name"] in self.TOOLTIP_KEEP_NAMES:
                self.create_tooltip(btn, item["add_tooltip"])

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.update_menu_items_status()

    # ---- item dispatch ----

    def toggle_menu_item(self, item):
        try:
            exists = item["check_command"]()
            logger.log(f"toggle '{item['name']}' (was {exists})", module=MOD)
            if exists:
                item["remove_command"]()
                item["enabled"] = False
            else:
                item["add_command"]()
                item["enabled"] = True
            self.update_indicator(item, not exists)
            self.update_menu_items_status()
            if item["name"] in ["Выключение/Перезагрузка/Гибернация",
                                "Диспетчер задач", "Перезапустить explorer",
                                "Очистить буфер", "Редактор реестра"]:
                menu_shell_order.update_desktop_shell_order()
        except Exception as e:
            logger.log(f"ERROR toggle_menu_item '{item['name']}': {e}", module=MOD)
            messagebox.showerror("Ошибка",
                                 f"Ошибка при переключении пункта: {e}")

    def set_default_settings(self):
        try:
            menu_checks.invalidate_print_cache()
            logger.log("set default context menu settings", module=MOD)
            for entry in self.menu_item_buttons:
                item = entry["data"]
                sc = item.get("supported_check")
                if sc is not None and not logger.is_debug() and not sc():
                    continue
                try:
                    cur = item["check_command"]()
                    if item.get("default_enabled", True) and not cur:
                        item["add_command"]()
                    elif not item.get("default_enabled", True) and cur:
                        item["remove_command"]()
                except Exception:
                    pass
            menu_power.update_power_menu_icons()
            self.update_menu_items_status()
            menu_shell_order.update_desktop_shell_order()
        except Exception as e:
            logger.log(f"ERROR set_default_settings: {e}", module=MOD)
            messagebox.showerror("Ошибка", f"Ошибка при сбросе настроек: {e}")

    def update_menu_items_status(self):
        try:
            for entry in self.menu_item_buttons:
                item = entry["data"]
                exists = item["check_command"]()
                item["enabled"] = exists
                self.update_indicator(item, exists)
                entry["button"].configure(text=item["name"])
        except Exception:
            pass

    def check_menu_item_exists(self, item_type):
        return menu_checks.check_menu_item_exists(item_type)

    def restart_explorer(self):
        if not win_utils.restart_explorer(MOD):
            messagebox.showerror("Ошибка", "Ошибка при перезагрузке проводника")

    # ---- UI helpers ----

    def create_menu_item_button(self, parent, item):
        btn_frame = ttk.Frame(parent)
        btn_frame.pack()

        indicator_canvas = tk.Canvas(btn_frame, width=20, height=20,
                                     highlightthickness=0)
        indicator_canvas.pack(side=tk.LEFT, padx=(0, 5))
        indicator_canvas.create_oval(2, 2, 18, 18, outline='#e74c3c',
                                     fill='#e74c3c')
        indicator_canvas.create_text(10, 10, text="✗", fill="white",
                                     font=('Arial', 10, 'bold'))

        item["indicator_canvas"] = indicator_canvas
        item["indicator_circle"] = indicator_canvas.find_all()[0]
        item["indicator_text"] = indicator_canvas.find_all()[1]

        btn = ttk.Button(
            btn_frame, text=item["name"], width=50,
            command=lambda i=item: self.toggle_menu_item(i))
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), pady=2)
        return btn

    def update_indicator(self, item, is_enabled):
        if 'indicator_canvas' in item and item['indicator_canvas']:
            canvas = item['indicator_canvas']
            if is_enabled:
                canvas.itemconfig(item['indicator_circle'],
                                  outline='#27ae60', fill='#27ae60')
                canvas.itemconfig(item['indicator_text'], text="✓")
            else:
                canvas.itemconfig(item['indicator_circle'],
                                  outline='#e74c3c', fill='#e74c3c')
                canvas.itemconfig(item['indicator_text'], text="✗")

    def create_tooltip(self, widget, text):
        return add_tooltip(widget, text, font=("Arial", 9), bg="#ffffe0")