import os
import tkinter as tk
import winreg
from tkinter import ttk, messagebox

HKCR = winreg.HKEY_CLASSES_ROOT
HKLM = winreg.HKEY_LOCAL_MACHINE
HKCU = winreg.HKEY_CURRENT_USER


def _reg_exists(hive, path):
    try:
        with winreg.OpenKey(hive, path):
            return True
    except (FileNotFoundError, OSError):
        return False


def _delete_key_recursive(hive, path):
    try:
        with winreg.OpenKey(hive, path) as k:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(k, i)
                except OSError:
                    break
                _delete_key_recursive(hive, path + "\\" + sub)
                i += 1
    except FileNotFoundError:
        return
    try:
        winreg.DeleteKey(hive, path)
    except OSError:
        pass


def _copy_key(hive, src, dst):
    winreg.CreateKey(hive, dst)
    with winreg.OpenKey(hive, src) as sk, winreg.OpenKey(hive, dst, 0, winreg.KEY_WRITE) as dk:
        i = 0
        while True:
            try:
                name, value, vtype = winreg.EnumValue(sk, i)
            except OSError:
                break
            winreg.SetValueEx(dk, name, 0, vtype, value)
            i += 1
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(sk, i)
            except OSError:
                break
            _copy_key(hive, src + "\\" + sub, dst + "\\" + sub)
            i += 1


def _rename_key(hive, old, new):
    _copy_key(hive, old, new)
    _delete_key_recursive(hive, old)


def _ensure_shellnew(hive, base_path):
    parent = '\\'.join(base_path.split('\\')[:-1])
    winreg.CreateKey(hive, parent)
    winreg.CreateKey(hive, base_path)
    with winreg.OpenKey(hive, base_path, 0, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, "NullFile", 0, winreg.REG_SZ, "")


FILE_NAMES = {
    '.folder': 'Папка', '.bmp': 'Рисунок BMP', '.contact': 'Контакт',
    '.lnk': 'Ярлык', '.rtf': 'Документ RTF', '.txt': 'Текстовый документ',
    '.zip': 'Архив ZIP',
}

TARGET_EXTENSIONS = {
    '.bmp': r'SOFTWARE\Classes\.bmp\ShellNew',
    '.contact': r'SOFTWARE\Classes\.contact\ShellNew',
    '.lnk': r'SOFTWARE\Classes\.lnk\ShellNew',
    '.rtf': r'SOFTWARE\Classes\.rtf\ShellNew',
    '.txt': r'SOFTWARE\Classes\.txt\ShellNew',
    '.zip': r'SOFTWARE\Classes\.zip\CompressedFolder\ShellNew',
}

# ---------- Работа с MS-расширениями и офисными клиентами ----------

MS_EXTENSIONS = {
    '.docx': 'Документ Word',
    '.xlsx': 'Книга Excel',
    '.pptx': 'Презентация PowerPoint',
    '.pub':  'Публикация Publisher',
}

MS_CLASSES = {
    '.docx': ['Word.Document', 'Word.DocumentMacroEnabled'],
    '.xlsx': ['Excel.Sheet', 'Excel.SheetMacroEnabled'],
    '.pptx': ['PowerPoint.Show', 'PowerPoint.ShowMacroEnabled'],
    '.pub':  ['Publisher.Document'],
}

THIRD_PARTY_CLASS_HINTS = [
    ('kwps', 'WPS Writer'),
    ('wps.document', 'WPS Writer'),
    ('ket.', 'WPS Spreadsheets'),
    ('wps.sheet', 'WPS Spreadsheets'),
    ('kwpp', 'WPS Presentation'),
    ('wps.presentation', 'WPS Presentation'),
    ('onlyoffice.document', 'OnlyOffice'),
    ('onlyoffice.sheet', 'OnlyOffice'),
    ('onlyoffice.slide', 'OnlyOffice'),
    ('libreoffice.writer', 'LibreOffice'),
    ('libreoffice.calc', 'LibreOffice'),
    ('libreoffice.impress', 'LibreOffice'),
    ('openoffice.writer', 'OpenOffice'),
    ('openoffice.calc', 'OpenOffice'),
    ('openoffice.impress', 'OpenOffice'),
    ('softmaker.textmaker', 'SoftMaker'),
    ('softmaker.planmaker', 'SoftMaker'),
    ('softmaker.presentations', 'SoftMaker'),
]


def _ext_registered(ext):
    try:
        with winreg.OpenKey(HKCR, ext):
            return True
    except (FileNotFoundError, OSError):
        return False


def _find_shellnew_under_ext(hive, ext):
    base = rf'SOFTWARE\Classes\{ext}'
    results = []
    try:
        with winreg.OpenKey(hive, base) as k:
            i = 0
            while True:
                try:
                    cls = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                cls_path = f'{base}\\{cls}'
                if _reg_exists(hive, cls_path + r'\ShellNew'):
                    results.append((hive, cls_path + r'\ShellNew', cls, True))
                elif _reg_exists(hive, cls_path + r'\ShellNew_no'):
                    results.append((hive, cls_path + r'\ShellNew_no', cls, False))
    except (FileNotFoundError, OSError):
        pass
    return results


def _classify_provider(class_name, ext):
    """Возвращает (provider_id, display_name, vendor).

    Для Microsoft vendor = '' (в UI не показываем).
    Для сторонних vendor = имя клиента.
    """
    low = class_name.lower()

    for base_cls in MS_CLASSES.get(ext, []):
        if low.startswith(base_cls.lower()):
            pid = base_cls.lower().split('.')[0] + '_' + ext.lstrip('.')
            return (pid, MS_EXTENSIONS[ext], '')

    for hint, vendor in THIRD_PARTY_CLASS_HINTS:
        if hint in low:
            pid = vendor.lower().replace(' ', '_') + '_' + ext.lstrip('.')
            return (pid, f"{MS_EXTENSIONS[ext]} ({vendor})", vendor)

    pid = 'unknown_' + ext.lstrip('.') + '_' + low.split('.')[-1]
    return (pid, f"{MS_EXTENSIONS[ext]} (другое)", 'other')


class CreateNewModule:
    def __init__(self, parent):
        self.parent = parent
        self.create_widgets()

    # ---- UI ----
    def create_widgets(self):
        container = ttk.Frame(self.parent)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)
        scrollable.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        ttk.Label(scrollable, text="Управление пунктами меню 'Создать'",
                  font=('Arial', 12, 'bold'), foreground='#2c3e50').pack(pady=10)

        ctrl = ttk.Frame(scrollable)
        ctrl.pack(pady=5)
        center = ttk.Frame(ctrl)
        center.pack(expand=True)
        self.toggle_all_btn = ttk.Button(center, text="", width=15,
                                         command=self.toggle_all_items)
        self.toggle_all_btn.pack(padx=5, pady=5)

        ttk.Separator(scrollable, orient='horizontal').pack(fill=tk.X, padx=10, pady=10)

        self.items_frame = ttk.Frame(scrollable)
        self.items_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.refresh()

    def refresh(self):
        for w in self.items_frame.winfo_children():
            w.destroy()

        items = self._get_items()
        if not items:
            ttk.Label(self.items_frame, text="Пункты меню 'Создать' не найдены",
                      foreground='#7f8c8d').pack(pady=20)
            return

        all_disabled = all(not it["enabled"] for it in items)
        self.toggle_all_btn.config(text="Включить всё" if all_disabled else "Отключить всё")

        header = ttk.Frame(self.items_frame)
        header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header, text="Название", font=('Arial', 9, 'bold'), width=30).pack(
            side=tk.LEFT, padx=5)
        ttk.Label(header, text="Статус", font=('Arial', 9, 'bold'), width=15).pack(
            side=tk.LEFT, padx=5)
        ttk.Label(header, text="Действие", font=('Arial', 9, 'bold'), width=15).pack(
            side=tk.LEFT)

        for item in items:
            self._create_row(item)

    def _create_row(self, item):
        row = ttk.Frame(self.items_frame, relief='groove', borderwidth=1)
        row.pack(fill=tk.X, pady=2, padx=5)

        ttk.Label(row, text=item["name"], width=29, anchor='w').pack(
            side=tk.LEFT, padx=5, pady=5)

        enabled = item["enabled"]
        ttk.Label(row, text="✅ Включен" if enabled else "❌ Отключен", width=15,
                  foreground='#27ae60' if enabled else '#e74c3c').pack(
            side=tk.LEFT, padx=5, pady=5)

        ttk.Button(row, text="Отключить" if enabled else "Включить", width=15,
                   command=lambda i=item: self.toggle_item(i)).pack(
            side=tk.LEFT, padx=5, pady=5)

    # ---- items discovery ----
    def _get_items(self):
        items = []
        try:
            # Folder
            if _reg_exists(HKCR, r'Folder\ShellNew'):
                items.append({"name": "Папка", "enabled": True,
                              "hive": HKCR, "base_path": r'Folder\ShellNew'})
            elif _reg_exists(HKCR, r'Folder\ShellNew_no'):
                items.append({"name": "Папка", "enabled": False,
                              "hive": HKCR, "base_path": r'Folder\ShellNew'})
            else:
                items.append({"name": "Папка", "enabled": False,
                              "hive": HKCR, "base_path": r'Folder\ShellNew'})

            for ext, path in TARGET_EXTENSIONS.items():
                name = FILE_NAMES.get(ext, f"Файл {ext.lstrip('.')}")
                if _reg_exists(HKLM, path):
                    items.append({"name": name, "enabled": True,
                                  "hive": HKLM, "base_path": path})
                elif _reg_exists(HKLM, path + '_no'):
                    items.append({"name": name, "enabled": False,
                                  "hive": HKLM, "base_path": path})
                else:
                    items.append({"name": name, "enabled": False,
                                  "hive": HKLM, "base_path": path})

            items.extend(self._get_office_items())
        except Exception as e:
            print(f"Ошибка: {e}")
        return items

    def _get_office_items(self):
        """Только реально установленные офисные клиенты, перехватившие MS-форматы.

        Перебирает подветки ShellNew под .docx/.xlsx/.pptx/.pub в HKLM и HKCU.
        Провайдер определяется по имени класса. Дубликаты (один provider_id + ext)
        схлопываются — приоритет HKLM.
        """
        items = []
        seen = set()

        for ext in MS_EXTENSIONS:
            if not _ext_registered(ext):
                continue
            for hive in (HKLM, HKCU):
                for h, path, cls, enabled in _find_shellnew_under_ext(hive, ext):
                    pid, display, _vendor = _classify_provider(cls, ext)
                    key = (pid, ext)
                    if key in seen:
                        continue
                    seen.add(key)
                    items.append({
                        "name": display,
                        "provider_id": pid,
                        "enabled": enabled,
                        "hive": h,
                        "shellnew": (h, path),
                    })
        return items

    # ---- actions ----
    def toggle_all_items(self):
        try:
            items = self._get_items()
            if not items:
                return
            all_disabled = all(not it["enabled"] for it in items)
            for item in items:
                if all_disabled and not item["enabled"]:
                    self._enable(item)
                elif not all_disabled and item["enabled"]:
                    self._disable(item)
            self.refresh()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении пунктов: {e}")

    def toggle_item(self, item):
        try:
            if item["enabled"]:
                self._disable(item)
            else:
                self._enable(item)
            self.refresh()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при переключении: {e}")

    def _enable(self, item):
        shellnew = item.get("shellnew")
        if shellnew:
            hive, path = shellnew
            try:
                if path.endswith('_no'):
                    _rename_key(hive, path, path[:-3])
                elif not _reg_exists(hive, path):
                    _ensure_shellnew(hive, os.path.dirname(path))
            except Exception:
                pass
            return
        hive, base = item["hive"], item["base_path"]
        try:
            if _reg_exists(hive, base + '_no'):
                _rename_key(hive, base + '_no', base)
            elif not _reg_exists(hive, base):
                _ensure_shellnew(hive, base)
        except Exception:
            pass

    def _disable(self, item):
        shellnew = item.get("shellnew")
        if shellnew:
            hive, path = shellnew
            try:
                if not path.endswith('_no') and _reg_exists(hive, path):
                    _rename_key(hive, path, path + '_no')
            except Exception:
                pass
            return
        hive, base = item["hive"], item["base_path"]
        try:
            if _reg_exists(hive, base):
                _rename_key(hive, base, base + '_no')
        except Exception:
            pass