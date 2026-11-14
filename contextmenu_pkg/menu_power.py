"""Меню питания: Гибернация / Перезагрузка / Выключение."""
import winreg

import logger
from core import reg_utils
from contextmenu_pkg.menu_helpers import get_imaster_path
from contextmenu_pkg.menu_registry_helpers import (
    take_ownership_registry_key_simple, set_registry_key_permissions_simple,
)

MOD = "contextmenu"


def update_power_menu_icons():
    imaster_dll_path = get_imaster_path()
    if not imaster_dll_path:
        for key_path in [r'DesktopBackground\Shell\Reboot',
                         r'DesktopBackground\Shell\Hibernation',
                         r'DesktopBackground\Shell\Clip']:
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0,
                                    winreg.KEY_SET_VALUE) as key:
                    try:
                        winreg.DeleteValue(key, "Icon")
                    except FileNotFoundError:
                        pass
            except Exception:
                pass
        return

    for key_path, idx in [(r'DesktopBackground\Shell\Reboot', -1001),
                          (r'DesktopBackground\Shell\Hibernation', -1002),
                          (r'DesktopBackground\Shell\Clip', -1003)]:
        try:
            icon = f"{imaster_dll_path},{idx}"
            if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                reg_utils.reg_add(key_path, 'Icon', icon, 'REG_EXPAND_SZ',
                                  module=MOD)
            else:
                reg_utils.reg_add(key_path, 'Icon', icon, 'REG_SZ', module=MOD)
        except Exception:
            pass


def add_power_menu():
    try:
        imaster_dll_path = get_imaster_path()

        hibernate_key = r'DesktopBackground\Shell\Hibernation'
        reg_utils.reg_add(hibernate_key, 'MUIVerb', 'Гибернация',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(hibernate_key, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(hibernate_key, 'Extended', '', 'REG_SZ', module=MOD)
        if imaster_dll_path:
            icon = f"{imaster_dll_path},-1002"
            if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                reg_utils.reg_add(hibernate_key, 'Icon', icon, 'REG_EXPAND_SZ',
                                  module=MOD)
            else:
                reg_utils.reg_add(hibernate_key, 'Icon', icon, 'REG_SZ', module=MOD)
        reg_utils.reg_add(hibernate_key + r'\command', '', "shutdown /h",
                          'REG_SZ', module=MOD)

        reboot_key = r'DesktopBackground\Shell\Reboot'
        reg_utils.reg_add(reboot_key, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(reboot_key, 'MUIVerb', 'Перезагрузка компьютера',
                          'REG_SZ', module=MOD)
        if imaster_dll_path:
            icon = f"{imaster_dll_path},-1001"
            if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                reg_utils.reg_add(reboot_key, 'Icon', icon, 'REG_EXPAND_SZ',
                                  module=MOD)
            else:
                reg_utils.reg_add(reboot_key, 'Icon', icon, 'REG_SZ', module=MOD)
        reg_utils.reg_add(reboot_key + r'\Command', '', "shutdown /r /t 0",
                          'REG_SZ', module=MOD)

        shutdown_key = r'DesktopBackground\Shell\Shutdown'
        reg_utils.reg_add(shutdown_key, 'Icon',
                          r"%SystemRoot%\System32\shell32.dll,-28",
                          'REG_EXPAND_SZ', module=MOD)
        reg_utils.reg_add(shutdown_key, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(shutdown_key, 'MUIVerb', 'Выключение компьютера',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(shutdown_key + r'\Command', '', "shutdown /s /t 0",
                          'REG_SZ', module=MOD)

        from contextmenu_pkg.menu_shell_order import update_desktop_shell_order
        update_desktop_shell_order()
    except PermissionError:
        if take_ownership_registry_key_simple(r'DesktopBackground\Shell'):
            if set_registry_key_permissions_simple(r'DesktopBackground\Shell'):
                try:
                    add_power_menu()
                except Exception:
                    from tkinter import messagebox
                    messagebox.showerror("Ошибка",
                                         "Не удалось добавить меню питания")
            else:
                from tkinter import messagebox
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            from tkinter import messagebox
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_power_menu: {e}", module=MOD)
        from tkinter import messagebox
        messagebox.showerror("Ошибка",
                             f"Ошибка при добавлении меню питания: {e}")


def remove_power_menu():
    try:
        for key_path in [r'DesktopBackground\Shell\Hibernation',
                         r'DesktopBackground\Shell\Reboot',
                         r'DesktopBackground\Shell\Shutdown']:
            try:
                reg_utils.reg_delete(key_path, module=MOD)
            except Exception:
                pass
        from contextmenu_pkg.menu_shell_order import update_desktop_shell_order
        update_desktop_shell_order()
    except Exception as e:
        logger.log(f"ERROR remove_power_menu: {e}", module=MOD)
        from tkinter import messagebox
        messagebox.showerror("Ошибка", f"Ошибка при удалении меню питания: {e}")