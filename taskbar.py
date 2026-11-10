import re
import subprocess
import sys
import time
import tkinter as tk
from tkinter import ttk, messagebox

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


def _run(args, **kwargs):
    kwargs.setdefault("creationflags", CREATE_NO_WINDOW)
    return subprocess.run(args, **kwargs)


# ---------- Модульные проверки версии/редакции (без UI) ----------

def get_windows_version():
    """Возвращает полную версию Windows в формате major.minor.build.revision."""
    try:
        r = _run(['wmic', 'os', 'get', 'Version', '/value'],
                 capture_output=True, text=True, encoding='cp866')
        m = re.search(r'Version=(\d+\.\d+\.\d+\.\d+)', r.stdout)
        if m:
            return m.group(1)
    except Exception:
        pass
    try:
        r = _run(['reg', 'query',
                  'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion',
                  '/v', 'CurrentBuildNumber'],
                 capture_output=True, text=True, encoding='cp866')
        m = re.search(r'CurrentBuildNumber\s+REG_SZ\s+(\d+)', r.stdout)
        if m:
            build = m.group(1)
            rev = 0
            try:
                r2 = _run(['reg', 'query',
                           'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion',
                           '/v', 'UBR'],
                          capture_output=True, text=True, encoding='cp866')
                ubr = re.search(r'UBR\s+REG_DWORD\s+0x([0-9a-fA-F]+)', r2.stdout)
                if ubr:
                    rev = int(ubr.group(1), 16)
            except Exception:
                pass
            return f"10.0.{build}.{rev}"
    except Exception:
        pass
    return "10.0.0.0"


def get_build_number(version=None):
    if version is None:
        version = get_windows_version()
    try:
        return int(version.split('.')[2])
    except Exception:
        return 0


def get_revision_number(version=None):
    if version is None:
        version = get_windows_version()
    try:
        parts = version.split('.')
        return int(parts[3]) if len(parts) > 3 else 0
    except Exception:
        return 0


def get_edition_id():
    """Возвращает EditionID из реестра (Enterprise, EnterpriseS, Professional и т.д.)."""
    try:
        r = _run(['reg', 'query',
                  'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion',
                  '/v', 'EditionID'],
                 capture_output=True, text=True, encoding='cp866')
        m = re.search(r'EditionID\s+REG_SZ\s+(\S+)', r.stdout)
        return m.group(1) if m else ''
    except Exception:
        return ''


def is_enterprise():
    """Корпоративные редакции: Enterprise или EnterpriseS (LTSC)."""
    return get_edition_id() in ('Enterprise', 'EnterpriseS')


def is_corporate_edition():
    """Для тултипов: Enterprise или EnterpriseS."""
    return is_enterprise()


def is_windows_version_supported():
    """Meet Now: 2004 (>=19041.610), 20H2 (>=19042.610), 21H2+ (>=19044).

    Не поддерживается в корпоративных редакциях (Enterprise, EnterpriseS).
    """
    if is_corporate_edition():
        return False
    build = get_build_number()
    rev = get_revision_number()
    if build == 19041:
        return rev >= 610
    if build == 19042:
        return rev >= 610
    return build >= 19044


def is_people_bar_supported():
    """Кнопка 'Люди' доступна в Windows 10 1709, 1803, 1809, 1903.

    Сквозная проверка по диапазону build:
    16299 (1709) <= build < 18363 (1909).
    Покрывает все накопительные обновления внутри каждой версии.
    В корпоративных редакциях (Enterprise, EnterpriseS) кнопка недоступна.
    """
    if is_corporate_edition():
        return False
    build = get_build_number()
    return 16299 <= build < 18363


class Module:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.windows_version = get_windows_version()
        self._create_widgets()
        self.show()

    # ---- тонкие обёртки для обратной совместимости ----
    def _get_build_number(self):
        return get_build_number(self.windows_version)

    def _get_revision_number(self):
        return get_revision_number(self.windows_version)

    def _get_edition_id(self):
        return get_edition_id()

    def is_enterprise(self):
        return is_enterprise()

    def is_corporate_edition(self):
        return is_corporate_edition()

    def is_windows_version_supported(self):
        return is_windows_version_supported()

    def is_people_bar_supported(self):
        return is_people_bar_supported()

    def _tooltip(self, widget, text):
        def enter(event):
            tip = tk.Toplevel()
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tk.Label(tip, text=text, background="lightyellow", relief="solid",
                     borderwidth=1, font=("Arial", 8)).pack()
            widget.tooltip = tip

        def leave(event):
            if getattr(widget, 'tooltip', None):
                widget.tooltip.destroy()
                widget.tooltip = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    # ---- state checks ----
    def check_weekday_state(self):
        r = _run(['reg', 'query', 'HKCU\\Control Panel\\International',
                  '/v', 'sShortDate'], capture_output=True, text=True, encoding='cp866')
        return "ddd" in r.stdout

    def check_notification_center_state(self):
        r = _run(['reg', 'query',
                  'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                  '/v', 'DisableNotificationCenter'],
                 capture_output=True, text=True, encoding='cp866')
        return "0x1" in r.stdout

    def check_taskbar_size_state(self):
        r = _run(['reg', 'query',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                  '/v', 'TaskbarSmallIcons'],
                 capture_output=True, text=True, encoding='cp866')
        return "0x1" in r.stdout

    def check_search_state(self):
        r = _run(['reg', 'query',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
                  '/v', 'SearchboxTaskbarMode'],
                 capture_output=True, text=True, encoding='cp866')
        if "0x0" in r.stdout:
            return 0
        if "0x1" in r.stdout:
            return 1
        return 2

    def check_people_state(self):
        r = _run(['reg', 'query',
                  'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                  '/v', 'HidePeopleBar'],
                 capture_output=True, text=True, encoding='cp866')
        return "0x1" in r.stdout

    def check_task_view_state(self):
        r = _run(['reg', 'query',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                  '/v', 'ShowTaskViewButton'],
                 capture_output=True, text=True, encoding='cp866')
        return "0x0" in r.stdout

    def check_meet_now_state(self):
        r = _run(['reg', 'query',
                  'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                  '/v', 'HideSCAMeetNow'],
                 capture_output=True, text=True, encoding='cp866')
        return "0x1" in r.stdout

    def check_grouping_state(self):
        r = _run(['reg', 'query',
                  'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
                  '/v', 'TaskbarGlomLevel'],
                 capture_output=True, text=True, encoding='cp866')
        if "0x0" in r.stdout:
            return 0
        if "0x1" in r.stdout:
            return 1
        if "0x2" in r.stdout:
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
            settings, text="Изменить группировку кнопок" + grouping_texts[grouping_state],
            command=self._toggle_grouping, width=40)
        self.toggle_grouping_button.pack(pady=5, fill=tk.X)
        self._tooltip(self.toggle_grouping_button, "Всегда → При заполнении → Никогда")

        size = self.check_taskbar_size_state()
        self.toggle_taskbar_size_button = ttk.Button(
            settings,
            text="Переключить размер значков панели задач"
                 + (" (Маленькие)" if size else " (Большие)"),
            command=self._toggle_taskbar_size, width=40)
        self.toggle_taskbar_size_button.pack(pady=5, fill=tk.X)
        self._tooltip(self.toggle_taskbar_size_button,
                      "Переключает между большими и маленькими значками")

        search = self.check_search_state()
        search_texts = {0: " (Скрыто)", 1: " (Значок)", 2: " (Поле)"}
        self.toggle_search_button = ttk.Button(
            settings, text="Скрыть/Показать Поиск" + search_texts[search],
            command=self._toggle_search, width=40)
        self.toggle_search_button.pack(pady=5, fill=tk.X)
        self._tooltip(self.toggle_search_button, "Поле → Значок → Скрыто")

        tv = self.check_task_view_state()
        self.toggle_task_view_button = ttk.Button(
            settings,
            text="Скрыть/Показать Просмотр задач"
                 + (" (Показано)" if not tv else " (Скрыто)"),
            command=self._toggle_task_view, width=40)
        self.toggle_task_view_button.pack(pady=5, fill=tk.X)
        self._tooltip(self.toggle_task_view_button,
                      "Скрывает или показывает кнопку 'Просмотр задач'")

        # --- People: Windows 10 1709 / 1803 / 1809 / 1903, не Enterprise/EnterpriseS ---
        people_supported = self.is_people_bar_supported()
        people = self.check_people_state() if people_supported else False
        people_text = (" (Показано)" if not people else " (Скрыто)") if people_supported else ""
        self.toggle_people_button = ttk.Button(
            settings,
            text="Скрыть/Показать Люди" + people_text,
            command=self._toggle_people, width=40)
        self.toggle_people_button.pack(pady=5, fill=tk.X)
        if people_supported:
            self._tooltip(self.toggle_people_button,
                          "Скрывает или показывает кнопку 'Люди'")
        else:
            self.toggle_people_button.config(state='disabled')
            if self.is_corporate_edition():
                tip_text = "Недоступно в Корпоративных редакциях"
            else:
                tip_text = "Доступно только в Windows 10 1709/1803/1809/1903"
            self._tooltip(self.toggle_people_button, tip_text)

        notif = self.check_notification_center_state()
        self.toggle_notification_center_button = ttk.Button(
            settings,
            text="Добавить/Убрать Центр уведомлений"
                 + (" (Показано)" if not notif else " (Скрыто)"),
            command=self._toggle_notification_center, width=40)
        self.toggle_notification_center_button.pack(pady=5, fill=tk.X)
        self._tooltip(self.toggle_notification_center_button,
                      "Включает или отключает Центр уведомлений")

        # --- Meet Now: 2004 (>=19041.610), 20H2 (>=19042.610), 21H2+ (>=19044), не Enterprise/EnterpriseS ---
        supported = self.is_windows_version_supported()
        meet = self.check_meet_now_state() if supported else False
        meet_text = (" (Показано)" if not meet else " (Скрыто)") if supported else ""
        self.toggle_meet_now_button = ttk.Button(
            settings,
            text="Скрыть/Показать Провести собрание" + meet_text,
            command=self._toggle_meet_now, width=40)
        self.toggle_meet_now_button.pack(pady=5, fill=tk.X)
        if supported:
            self._tooltip(self.toggle_meet_now_button,
                          "Скрывает или показывает кнопку 'Провести собрание'")
        else:
            self.toggle_meet_now_button.config(state='disabled')
            if self.is_corporate_edition():
                tip_text = "Недоступно в Корпоративных редакциях"
            else:
                tip_text = ("Доступно только в Windows 10 2004/20H2 "
                            "(с KB4580364) и новее")
            self._tooltip(self.toggle_meet_now_button, tip_text)

        weekday = self.check_weekday_state()
        self.toggle_weekday_button = ttk.Button(
            settings,
            text="Добавить/Убрать день недели в дате"
                 + (" (Показано)" if weekday else " (Скрыто)"),
            command=self._toggle_weekday, width=40)
        self.toggle_weekday_button.pack(pady=5, fill=tk.X)
        self._tooltip(self.toggle_weekday_button,
                      "Добавляет или убирает день недели в системной дате")

        extra = ttk.Frame(settings)
        extra.pack(fill=tk.X, pady=10)
        restart_btn = ttk.Button(extra, text="Перезапустить Проводник",
                                 command=self._restart_explorer, width=18)
        restart_btn.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        self._tooltip(restart_btn, "Перезапускает Проводник Windows")
        reset_btn = ttk.Button(extra, text="Сбросить все настройки",
                               command=self._reset_all, width=18)
        reset_btn.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        self._tooltip(reset_btn, "Сбрасывает все настройки панели задач")

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
        if not self.is_people_bar_supported():
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
        if not self.is_windows_version_supported():
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
        r = _run(['reg', 'query', 'HKCU\\Control Panel\\International',
                  '/v', 'sShortDate'], capture_output=True, text=True, encoding='cp866')
        value = 'dd.MM.yyyy' if "ddd" in r.stdout else 'ddd dd.MM.yyyy'
        _run(['reg', 'add', 'HKCU\\Control Panel\\International',
              '/V', 'sShortDate', '/D', value, '/T', 'REG_SZ', '/f'],
             capture_output=True)

    def _do_toggle_notification_center(self):
        current = 1 if self.check_notification_center_state() else 0
        if current:
            try:
                _run(['reg', 'add',
                      'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                      '/V', 'DisableNotificationCenter', '/D', '0',
                      '/T', 'REG_DWORD', '/f'], capture_output=True, check=True)
            except subprocess.CalledProcessError:
                _run(['reg', 'delete',
                      'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                      '/v', 'DisableNotificationCenter', '/f'], capture_output=True)
        else:
            _run(['reg', 'add',
                  'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer',
                  '/V', 'DisableNotificationCenter', '/D', '1',
                  '/T', 'REG_DWORD', '/f'], capture_output=True)

    def _do_toggle_taskbar_size(self):
        new = '0' if self.check_taskbar_size_state() else '1'
        _run(['reg', 'add',
              'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
              '/V', 'TaskbarSmallIcons', '/D', new, '/T', 'REG_DWORD', '/f'],
             capture_output=True)

    def _do_toggle_search(self):
        current = self.check_search_state()
        new = {2: 1, 1: 0, 0: 2}[current]
        _run(['reg', 'add',
              'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search',
              '/V', 'SearchboxTaskbarMode', '/D', str(new),
              '/T', 'REG_DWORD', '/f'], capture_output=True)

    def _do_toggle_people(self):
        if self.check_people_state():
            try:
                _run(['reg', 'delete',
                      'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                      '/v', 'HidePeopleBar', '/f'], capture_output=True, check=True)
            except subprocess.CalledProcessError:
                _run(['reg', 'add',
                      'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                      '/v', 'HidePeopleBar', '/d', '0', '/t', 'REG_DWORD', '/f'],
                     capture_output=True)
        else:
            _run(['reg', 'add',
                  'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                  '/v', 'HidePeopleBar', '/d', '1', '/t', 'REG_DWORD', '/f'],
                 capture_output=True)

    def _do_toggle_task_view(self):
        new = '1' if self.check_task_view_state() else '0'
        _run(['reg', 'add',
              'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
              '/V', 'ShowTaskViewButton', '/D', new, '/T', 'REG_DWORD', '/f'],
             capture_output=True)

    def _do_toggle_meet_now(self):
        if self.check_meet_now_state():
            try:
                _run(['reg', 'delete',
                      'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                      '/v', 'HideSCAMeetNow', '/f'], capture_output=True, check=True)
            except subprocess.CalledProcessError:
                _run(['reg', 'add',
                      'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                      '/v', 'HideSCAMeetNow', '/d', '0', '/t', 'REG_DWORD', '/f'],
                     capture_output=True)
        else:
            _run(['reg', 'add',
                  'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                  '/v', 'HideSCAMeetNow', '/d', '1', '/t', 'REG_DWORD', '/f'],
                 capture_output=True)

    def _do_toggle_grouping(self):
        current = self.check_grouping_state()
        new = {0: 1, 1: 2, 2: 0}[current]
        _run(['reg', 'add',
              'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced',
              '/V', 'TaskbarGlomLevel', '/D', str(new), '/T', 'REG_DWORD', '/f'],
             capture_output=True)

    def _restart_explorer(self):
        try:
            _run(['taskkill', '/f', '/im', 'explorer.exe'], capture_output=True)
            time.sleep(1)
            subprocess.Popen(['explorer.exe'], creationflags=CREATE_NO_WINDOW)
        except Exception:
            try:
                subprocess.Popen(['explorer.exe'])
            except Exception:
                pass

    def _reset_all(self):
        try:
            settings = [
                ('HKCU\\Control Panel\\International', 'sShortDate', 'dd.MM.yyyy', 'REG_SZ'),
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
            for key, name, value, vtype in settings:
                _run(['reg', 'add', key, '/V', name, '/D', value, '/T', vtype, '/f'],
                     capture_output=True)

            if self.is_people_bar_supported():
                _run(['reg', 'delete',
                      'HKCU\\Software\\Policies\\Microsoft\\Windows\\Explorer',
                      '/v', 'HidePeopleBar', '/f'], capture_output=True)
            if self.is_windows_version_supported():
                _run(['reg', 'delete',
                      'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer',
                      '/v', 'HideSCAMeetNow', '/f'], capture_output=True)

            self._update_all_buttons()
            self._restart_explorer()
        except Exception as e:
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