import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

import logger
from core import win_utils, reg_utils
from core.ui_utils import add_tooltip

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
MOD = "taskbar"


# ---------- Тонкие обёртки для обратной совместимости ----------
# Внешний код (about.py) вызывает taskbar.is_people_bar_supported() и
# taskbar.is_windows_version_supported() — оставляем как прокси в core.win_utils.

def is_windows_version_supported():
    return win_utils.is_meet_now_supported()


def is_people_bar_supported():
    return win_utils.is_people_bar_supported()


def is_enterprise():
    return win_utils.is_corporate_edition()


def is_corporate_edition():
    return win_utils.is_corporate_edition()


def get_windows_version():
    return win_utils.get_windows_version()


def get_build_number(version=None):
    return win_utils.get_build_number()


def get_revision_number(version=None):
    return win_utils.get_revision_number()


def get_edition_id():
    return win_utils.get_edition_id()


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.windows_version = win_utils.get_windows_version()
        self._create_widgets()
        self.show()

    # ---- тонкие обёртки для обратной совместимости ----
    def _get_build_number(self):
        return win_utils.get_build_number()

    def _get_revision_number(self):
        return win_utils.get_revision_number()

    def _get_edition_id(self):
        return win_utils.get_edition_id()

    def is_enterprise(self):
        return win_utils.is_corporate_edition()

    def is_corporate_edition(self):
        return win_utils.is_corporate_edition()

    def is_windows_version_supported(self):
        return win_utils.is_meet_now_supported()

    def is_people_bar_supported(self):
        return win_utils.is_people_bar_supported()

    # ---- state checks ----
    def check_weekday_state(self):
        r = reg_utils.reg('query', 'HKCU\\Control Panel\\International',
                          '/v', 'sShortDate', log=False)
        return "ddd" in (r.stdout or '')

    def check_notification_center_state(self):
        r = reg_utils.reg('query',
                          'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                          '/v', 'DisableNotificationCenter', log=False)
        return "0x1" in (r.stdout or '')

    def check_taskbar_size_state(self):
        r = reg_utils.reg(
            'query',
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
            '/v', 'TaskbarSmallIcons', log=False)
        return "0x1" in (r.stdout or '')

    def check_search_state(self):
        r = reg_utils.reg('query',
                          'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
                          '/v', 'SearchboxTaskbarMode', log=False)
        out = r.stdout or ''
        if "0x0" in out:
            return 0
        if "0x1" in out:
            return 1
        return 2

    def check_people_state(self):
        r = reg_utils.reg('query',
                          'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                          '/v', 'HidePeopleBar', log=False)
        return "0x1" in (r.stdout or '')

    def check_task_view_state(self):
        r = reg_utils.reg(
            'query',
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
            '/v', 'ShowTaskViewButton', log=False)
        return "0x0" in (r.stdout or '')

    def check_meet_now_state(self):
        r = reg_utils.reg(
            'query',
            'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
            '/v', 'HideSCAMeetNow', log=False)
        return "0x1" in (r.stdout or '')

    def check_grouping_state(self):
        r = reg_utils.reg(
            'query',
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
            '/v', 'TaskbarGlomLevel', log=False)
        out = r.stdout or ''
        if "0x0" in out:
            return 0
        if "0x1" in out:
            return 1
        if "0x2" in out:
            return 2
        return 0

    # ---- helpers ----
    def _update_button_text(self, button, base, state, is_search=False, is_grouping=False):
        if button == self.toggle_taskbar_size_button:
            button.config(text=f"{base} ({'Маленькие' if state else 'Большие'})")
        elif is_search:
            labels = {0: "Скрыто", 1: "Значок", 2: "Поле"}
            button.config(text=f"{base} ({labels.get(state, 'Поле')})")
        elif is_grouping:
            labels = {0: "Всегда", 1: "При заполнении", 2: "Никогда"}
            button.config(text=f"{base} ({labels.get(state, 'Всегда')})")
        else:
            button.config(text=f"{base} ({'Показано' if state else 'Скрыто'})")

    # ---- UI ----
    def _create_widgets(self):
        ttk.Label(self.frame, text="Настройка Панели задач",
                  font=('Arial', 14, 'bold'), foreground='#2c3e50').pack(pady=5)
        ttk.Label(self.frame,
                  text="Настройте внешний вид и функциональность Панели задач Windows",
                  font=('Arial', 9), foreground='#7f8c8d').pack(pady=(0, 10))

        settings = ttk.LabelFrame(self.frame, text="Доступные настройки", padding=10)
        settings.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        grouping_state = self.check_grouping_state()
        grouping_texts = {0: " (Всегда)", 1: " (При заполнении)", 2: " (Никогда)"}
        self.toggle_grouping_button = ttk.Button(
            settings,
            text="Изменить группировку кнопок" + grouping_texts[grouping_state],
            command=self._toggle_grouping, width=40)
        self.toggle_grouping_button.pack(pady=5, fill=tk.X)
        add_tooltip(self.toggle_grouping_button, "Всегда → При заполнении → Никогда")

        size = self.check_taskbar_size_state()
        self.toggle_taskbar_size_button = ttk.Button(
            settings,
            text="Переключить размер значков панели задач"
                 + (" (Маленькие)" if size else " (Большие)"),
            command=self._toggle_taskbar_size, width=40)
        self.toggle_taskbar_size_button.pack(pady=5, fill=tk.X)
        add_tooltip(self.toggle_taskbar_size_button,
                    "Переключает между большими и маленькими значками")

        search = self.check_search_state()
        search_texts = {0: " (Скрыто)", 1: " (Значок)", 2: " (Поле)"}
        self.toggle_search_button = ttk.Button(
            settings, text="Скрыть/Показать Поиск" + search_texts[search],
            command=self._toggle_search, width=40)
        self.toggle_search_button.pack(pady=5, fill=tk.X)
        add_tooltip(self.toggle_search_button, "Поле → Значок → Скрыто")

        tv = self.check_task_view_state()
        self.toggle_task_view_button = ttk.Button(
            settings,
            text="Скрыть/Показать Просмотр задач"
                 + (" (Показано)" if not tv else " (Скрыто)"),
            command=self._toggle_task_view, width=40)
        self.toggle_task_view_button.pack(pady=5, fill=tk.X)
        add_tooltip(self.toggle_task_view_button,
                    "Скрывает или показывает кнопку 'Просмотр задач'")

        # --- People ---
        people_supported = self.is_people_bar_supported()
        people = self.check_people_state() if people_supported else False
        people_text = (" (Показано)" if not people else " (Скрыто)") if people_supported else ""
        self.toggle_people_button = ttk.Button(
            settings,
            text="Скрыть/Показать Люди" + people_text,
            command=self._toggle_people, width=40)
        self.toggle_people_button.pack(pady=5, fill=tk.X)
        if people_supported:
            add_tooltip(self.toggle_people_button,
                        "Скрывает или показывает кнопку 'Люди'")
        elif not logger.is_debug():
            self.toggle_people_button.config(state='disabled')
            if self.is_corporate_edition():
                tip_text = "Недоступно в Корпоративных редакциях"
            else:
                tip_text = "Доступно только в Windows 10 1709/1803/1809/1903"
            add_tooltip(self.toggle_people_button, tip_text)

        notif = self.check_notification_center_state()
        self.toggle_notification_center_button = ttk.Button(
            settings,
            text="Добавить/Убрать Центр уведомлений"
                 + (" (Показано)" if not notif else " (Скрыто)"),
            command=self._toggle_notification_center, width=40)
        self.toggle_notification_center_button.pack(pady=5, fill=tk.X)
        add_tooltip(self.toggle_notification_center_button,
                    "Включает или отключает Центр уведомлений")

        # --- Meet Now ---
        supported = self.is_windows_version_supported()
        meet = self.check_meet_now_state() if supported else False
        meet_text = (" (Показано)" if not meet else " (Скрыто)") if supported else ""
        self.toggle_meet_now_button = ttk.Button(
            settings,
            text="Скрыть/Показать Провести собрание" + meet_text,
            command=self._toggle_meet_now, width=40)
        self.toggle_meet_now_button.pack(pady=5, fill=tk.X)
        if supported:
            add_tooltip(self.toggle_meet_now_button,
                        "Скрывает или показывает кнопку 'Провести собрание'")
        elif not logger.is_debug():
            self.toggle_meet_now_button.config(state='disabled')
            if self.is_corporate_edition():
                tip_text = "Недоступно в Корпоративных редакциях"
            else:
                tip_text = ("Доступно только в Windows 10 2004/20H2 "
                            "(с KB4580364) и новее")
            add_tooltip(self.toggle_meet_now_button, tip_text)

        weekday = self.check_weekday_state()
        self.toggle_weekday_button = ttk.Button(
            settings,
            text="Добавить/Убрать день недели в дате"
                 + (" (Показано)" if weekday else " (Скрыто)"),
            command=self._toggle_weekday, width=40)
        self.toggle_weekday_button.pack(pady=5, fill=tk.X)
        add_tooltip(self.toggle_weekday_button,
                    "Добавляет или убирает день недели в системной дате")

        extra = ttk.Frame(settings)
        extra.pack(fill=tk.X, pady=10)
        restart_btn = ttk.Button(extra, text="Перезапустить Проводник",
                                 command=self._restart_explorer, width=18)
        restart_btn.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        add_tooltip(restart_btn, "Перезапускает Проводник Windows")
        reset_btn = ttk.Button(extra, text="Сбросить все настройки",
                               command=self._reset_all, width=18)
        reset_btn.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        add_tooltip(reset_btn, "Сбрасывает все настройки панели задач")

    # ---- toggle wrappers ----
    def _toggle_weekday(self):
        self._do_toggle_weekday()
        self._update_button_text(self.toggle_weekday_button,
                                 "Добавить/Убрать день недели в дате",
                                 self.check_weekday_state())

    def _toggle_notification_center(self):
        self._do_toggle_notification_center()
        self._update_button_text(self.toggle_notification_center_button,
                                 "Добавить/Убрать Центр уведомлений",
                                 not self.check_notification_center_state())

    def _toggle_taskbar_size(self):
        self._do_toggle_taskbar_size()
        self._update_button_text(self.toggle_taskbar_size_button,
                                 "Переключить размер значков панели задач",
                                 self.check_taskbar_size_state())

    def _toggle_search(self):
        self._do_toggle_search()
        self._update_button_text(self.toggle_search_button, "Скрыть/Показать Поиск",
                                 self.check_search_state(), is_search=True)

    def _toggle_people(self):
        if not self.is_people_bar_supported() and not logger.is_debug():
            return
        self._do_toggle_people()
        self._update_button_text(self.toggle_people_button, "Скрыть/Показать Люди",
                                 not self.check_people_state())

    def _toggle_task_view(self):
        self._do_toggle_task_view()
        self._update_button_text(self.toggle_task_view_button,
                                 "Скрыть/Показать Просмотр задач",
                                 not self.check_task_view_state())

    def _toggle_meet_now(self):
        if not self.is_windows_version_supported() and not logger.is_debug():
            return
        self._do_toggle_meet_now()
        self._update_button_text(self.toggle_meet_now_button,
                                 "Скрыть/Показать Провести собрание",
                                 not self.check_meet_now_state())

    def _toggle_grouping(self):
        self._do_toggle_grouping()
        self._update_button_text(self.toggle_grouping_button,
                                 "Изменить группировку кнопок",
                                 self.check_grouping_state(), is_grouping=True)

    # ---- actual registry operations ----
    def _do_toggle_weekday(self):
        r = reg_utils.reg('query', 'HKCU\\Control Panel\\International',
                          '/v', 'sShortDate', log=False)
        value = 'dd.MM.yyyy' if "ddd" in (r.stdout or '') else 'ddd dd.MM.yyyy'
        reg_utils.reg_add('HKCU\\Control Panel\\International', 'sShortDate',
                          value, 'REG_SZ', module=MOD)

    def _do_toggle_notification_center(self):
        current = 1 if self.check_notification_center_state() else 0
        key = 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer'
        if current:
            rr = reg_utils.reg_add(key, 'DisableNotificationCenter',
                                   '0', 'REG_DWORD', module=MOD)
            if rr.returncode != 0:
                reg_utils.reg_delete(key, 'DisableNotificationCenter', module=MOD)
        else:
            reg_utils.reg_add(key, 'DisableNotificationCenter',
                              '1', 'REG_DWORD', module=MOD)

    def _do_toggle_taskbar_size(self):
        new = '0' if self.check_taskbar_size_state() else '1'
        reg_utils.reg_add(
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
            'TaskbarSmallIcons', new, 'REG_DWORD', module=MOD)

    def _do_toggle_search(self):
        current = self.check_search_state()
        new = {2: 1, 1: 0, 0: 2}[current]
        reg_utils.reg_add(
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
            'SearchboxTaskbarMode', str(new), 'REG_DWORD', module=MOD)

    def _do_toggle_people(self):
        key = 'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer'
        if self.check_people_state():
            rr = reg_utils.reg_delete(key, 'HidePeopleBar', module=MOD)
            if rr.returncode != 0:
                reg_utils.reg_add(key, 'HidePeopleBar', '0', 'REG_DWORD', module=MOD)
        else:
            reg_utils.reg_add(key, 'HidePeopleBar', '1', 'REG_DWORD', module=MOD)

    def _do_toggle_task_view(self):
        new = '1' if self.check_task_view_state() else '0'
        reg_utils.reg_add(
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
            'ShowTaskViewButton', new, 'REG_DWORD', module=MOD)

    def _do_toggle_meet_now(self):
        key = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer'
        if self.check_meet_now_state():
            rr = reg_utils.reg_delete(key, 'HideSCAMeetNow', module=MOD)
            if rr.returncode != 0:
                reg_utils.reg_add(key, 'HideSCAMeetNow', '0', 'REG_DWORD', module=MOD)
        else:
            reg_utils.reg_add(key, 'HideSCAMeetNow', '1', 'REG_DWORD', module=MOD)

    def _do_toggle_grouping(self):
        current = self.check_grouping_state()
        new = {0: 1, 1: 2, 2: 0}[current]
        reg_utils.reg_add(
            'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
            'TaskbarGlomLevel', str(new), 'REG_DWORD', module=MOD)

    def _restart_explorer(self):
        win_utils.restart_explorer(MOD)

    def _reset_all(self):
        try:
            settings = [
                ('HKCU\\Control Panel\\International', 'sShortDate',
                 'dd.MM.yyyy', 'REG_SZ'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                 'TaskbarSmallIcons', '0', 'REG_DWORD'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
                 'SearchboxTaskbarMode', '2', 'REG_DWORD'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                 'ShowTaskViewButton', '1', 'REG_DWORD'),
                ('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                 'TaskbarGlomLevel', '0', 'REG_DWORD'),
                ('HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                 'DisableNotificationCenter', '0', 'REG_DWORD'),
            ]
            logger.log("reset all taskbar settings", module=MOD)
            for key, name, value, vtype in settings:
                reg_utils.reg_add(key, name, value, vtype, module=MOD)

            if self.is_people_bar_supported():
                reg_utils.reg_delete(
                    'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                    'HidePeopleBar', module=MOD)
            if self.is_windows_version_supported():
                reg_utils.reg_delete(
                    'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                    'HideSCAMeetNow', module=MOD)

            self._update_all_buttons()
            self._restart_explorer()
        except Exception as e:
            logger.log(f"ERROR _reset_all: {e}", module=MOD)
            messagebox.showerror("Ошибка сброса настроек",
                                 f"Не удалось сбросить настройки\nОшибка: {e}")

    def _update_all_buttons(self):
        self._update_button_text(self.toggle_grouping_button,
                                 "Изменить группировку кнопок",
                                 self.check_grouping_state(), is_grouping=True)
        self._update_button_text(self.toggle_taskbar_size_button,
                                 "Переключить размер значков панели задач",
                                 self.check_taskbar_size_state())
        self._update_button_text(self.toggle_search_button, "Скрыть/Показать Поиск",
                                 self.check_search_state(), is_search=True)
        self._update_button_text(self.toggle_task_view_button,
                                 "Скрыть/Показать Просмотр задач",
                                 not self.check_task_view_state())
        if self.is_people_bar_supported():
            self._update_button_text(self.toggle_people_button, "Скрыть/Показать Люди",
                                     not self.check_people_state())
        self._update_button_text(self.toggle_notification_center_button,
                                 "Добавить/Убрать Центр уведомлений",
                                 not self.check_notification_center_state())
        if self.is_windows_version_supported():
            self._update_button_text(self.toggle_meet_now_button,
                                     "Скрыть/Показать Провести собрание",
                                     not self.check_meet_now_state())
        self._update_button_text(self.toggle_weekday_button,
                                 "Добавить/Убрать день недели в дате",
                                 self.check_weekday_state())

    def show(self):
        self.frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def hide(self):
        self.frame.pack_forget()