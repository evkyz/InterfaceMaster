"""Простые пункты контекстного меню: regedit, taskmanager, clipboard,
copy_as_path, pin_to_start/taskbar, paint3d, restart_explorer."""
import winreg
from tkinter import messagebox

import logger
from core import reg_utils, win_utils
from contextmenu_pkg.menu_helpers import get_imaster_path
from contextmenu_pkg.menu_shell_order import update_desktop_shell_order
from contextmenu_pkg.menu_registry_helpers import (
    take_ownership_registry_key_simple, set_registry_key_permissions_simple,
    set_registry_key_permissions, take_ownership_registry_key,
    delete_registry_key_recursive,
)
from contextmenu_pkg.constants import PAINT3D_EXTS

MOD = "contextmenu"


# ---- regedit ----

def add_regedit():
    try:
        key_path = r'DesktopBackground\Shell\Regedit'
        reg_utils.reg_add(key_path, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'MUIVerb', 'Редактор реестра',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Icon', 'regedit.exe,0', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Extended', '', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path + r'\command', '', 'regedit',
                          'REG_SZ', module=MOD)
        update_desktop_shell_order()
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    add_regedit()
                except Exception:
                    messagebox.showerror(
                        "Ошибка", "Не удалось добавить пункт 'Редактор реестра'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_regedit: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_regedit():
    try:
        key_path = r'DesktopBackground\Shell\Regedit'
        reg_utils.reg_delete(key_path, module=MOD)
        update_desktop_shell_order()
    except FileNotFoundError:
        pass
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                    update_desktop_shell_order()
                except Exception:
                    messagebox.showerror(
                        "Ошибка", "Не удалось удалить пункт 'Редактор реестра'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR remove_regedit: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- restart explorer ----

def add_restart_explorer():
    try:
        key_path = r'DesktopBackground\Shell\RestartExplorer'
        reg_utils.reg_add(key_path, 'Icon', 'explorer.exe', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'MUIVerb', 'Перезапустить explorer',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Extended', '', 'REG_SZ', module=MOD)
        reg_utils.reg_add(
            key_path + r'\command', '',
            "cmd.exe /c taskkill /f /im explorer.exe & start explorer.exe",
            'REG_SZ', module=MOD)
        update_desktop_shell_order()
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    add_restart_explorer()
                except Exception:
                    messagebox.showerror(
                        "Ошибка",
                        "Не удалось добавить пункт 'Перезапустить explorer'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_restart_explorer: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_restart_explorer():
    try:
        key_path = r'DesktopBackground\Shell\RestartExplorer'
        reg_utils.reg_delete(key_path, module=MOD)
        update_desktop_shell_order()
    except FileNotFoundError:
        pass
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                    update_desktop_shell_order()
                except Exception:
                    messagebox.showerror(
                        "Ошибка",
                        "Не удалось удалить пункт 'Перезапустить explorer'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR remove_restart_explorer: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- clipboard clear ----

def add_clipboard_clear():
    try:
        key_path = r'DesktopBackground\Shell\Clip'
        imaster_dll_path = get_imaster_path()
        reg_utils.reg_add(key_path, 'MUIVerb', 'Очистить буфер',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Extended', '', 'REG_SZ', module=MOD)
        if imaster_dll_path:
            icon = f"{imaster_dll_path},-1003"
            if "System32" in imaster_dll_path or "%SystemRoot%" in imaster_dll_path:
                reg_utils.reg_add(key_path, 'Icon', icon, 'REG_EXPAND_SZ',
                                  module=MOD)
            else:
                reg_utils.reg_add(key_path, 'Icon', icon, 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path + r'\command', '', "cmd /c echo off | clip",
                          'REG_SZ', module=MOD)
        update_desktop_shell_order()
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    add_clipboard_clear()
                except Exception:
                    messagebox.showerror(
                        "Ошибка", "Не удалось добавить пункт 'Очистить буфер'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_clipboard_clear: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_clipboard_clear():
    try:
        key_path = r'DesktopBackground\Shell\Clip'
        reg_utils.reg_delete(key_path, module=MOD)
        update_desktop_shell_order()
    except FileNotFoundError:
        pass
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                    update_desktop_shell_order()
                except Exception:
                    messagebox.showerror(
                        "Ошибка", "Не удалось удалить пункт 'Очистить буфер'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR remove_clipboard_clear: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- task manager ----

def add_task_manager():
    try:
        key_path = r'DesktopBackground\Shell\TaskManager'
        reg_utils.reg_add(key_path, 'Icon', 'taskmgr.exe,0', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, '', '', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Position', 'Bottom', 'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'MUIVerb', 'Диспетчер задач',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path + r'\Command', '', 'taskmgr.exe',
                          'REG_SZ', module=MOD)
        update_desktop_shell_order()
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    add_task_manager()
                except Exception:
                    messagebox.showerror(
                        "Ошибка", "Не удалось добавить пункт 'Диспетчер задач'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_task_manager: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_task_manager():
    try:
        key_path = r'DesktopBackground\Shell\TaskManager'
        reg_utils.reg_delete(key_path, module=MOD)
        update_desktop_shell_order()
    except FileNotFoundError:
        pass
    except PermissionError:
        if take_ownership_registry_key_simple(key_path):
            if set_registry_key_permissions_simple(key_path):
                try:
                    delete_registry_key_recursive(winreg.HKEY_CLASSES_ROOT, key_path)
                    update_desktop_shell_order()
                except Exception:
                    messagebox.showerror(
                        "Ошибка", "Не удалось удалить пункт 'Диспетчер задач'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR remove_task_manager: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- copy as path ----

def add_copy_as_path():
    try:
        key_path = (r'SOFTWARE\Classes\AllFilesystemObjects\shellex'
                    r'\ContextMenuHandlers\CopyAsPathMenu')
        reg_utils.reg_add(key_path, '', '{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}',
                          'REG_SZ', module=MOD)
    except PermissionError:
        if take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
            if set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                try:
                    add_copy_as_path()
                except Exception:
                    messagebox.showerror(
                        "Ошибка",
                        "Не удалось добавить пункт 'Копировать как путь'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_copy_as_path: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_copy_as_path():
    try:
        key_path = (r'SOFTWARE\Classes\AllFilesystemObjects\shellex'
                    r'\ContextMenuHandlers\CopyAsPathMenu')
        reg_utils.reg_delete(key_path, module=MOD)
    except FileNotFoundError:
        pass
    except PermissionError:
        if take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
            if set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                try:
                    delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE,
                                                  key_path)
                except Exception:
                    messagebox.showerror(
                        "Ошибка",
                        "Не удалось удалить пункт 'Копировать как путь'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR remove_copy_as_path: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- pin to start / taskbar ----

def add_pin_to_start():
    paths = (
        r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen',
        r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\PintoStartScreen',
    )
    ok = 0
    for key_path in paths:
        try:
            reg_utils.reg_add(key_path, '',
                              '{470C0EBD-5D73-4d58-9CED-E91E22E23282}',
                              'REG_SZ', module=MOD)
            ok += 1
        except PermissionError:
            if take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE,
                                                key_path):
                    try:
                        reg_utils.reg_add(
                            key_path, '',
                            '{470C0EBD-5D73-4d58-9CED-E91E22E23282}',
                            'REG_SZ', module=MOD)
                        ok += 1
                    except Exception:
                        pass
        except Exception:
            pass
    if ok == 0:
        messagebox.showerror(
            "Ошибка",
            "Не удалось добавить пункт 'Закрепить на начальном экране'")


def remove_pin_to_start():
    paths = (
        r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen',
        r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\PintoStartScreen',
    )
    for key_path in paths:
        try:
            reg_utils.reg_delete(key_path, module=MOD)
        except FileNotFoundError:
            pass
        except PermissionError:
            if take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
                if set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE,
                                                key_path):
                    try:
                        delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE,
                                                      key_path)
                    except Exception:
                        pass
        except Exception:
            pass


def add_pin_to_taskbar():
    try:
        key_path = (r'SOFTWARE\Classes\*\shellex'
                    r'\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}')
        reg_utils.reg_add(key_path, '', 'Taskband Pin', 'REG_SZ', module=MOD)
    except PermissionError:
        if take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
            if set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                try:
                    add_pin_to_taskbar()
                except Exception:
                    messagebox.showerror(
                        "Ошибка",
                        "Не удалось добавить пункт 'Закрепить на Панели задач'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR add_pin_to_taskbar: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_pin_to_taskbar():
    try:
        key_path = (r'SOFTWARE\Classes\*\shellex'
                    r'\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}')
        reg_utils.reg_delete(key_path, module=MOD)
    except FileNotFoundError:
        pass
    except PermissionError:
        if take_ownership_registry_key(winreg.HKEY_LOCAL_MACHINE, key_path):
            if set_registry_key_permissions(winreg.HKEY_LOCAL_MACHINE, key_path):
                try:
                    delete_registry_key_recursive(winreg.HKEY_LOCAL_MACHINE,
                                                  key_path)
                except Exception:
                    messagebox.showerror(
                        "Ошибка",
                        "Не удалось удалить пункт 'Закрепить на Панели задач'")
            else:
                messagebox.showerror("Ошибка", "Не удалось установить права доступа")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось получить владение ключом реестра")
    except Exception as e:
        logger.log(f"ERROR remove_pin_to_taskbar: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- Paint 3D ----

def add_paint3d_edit():
    try:
        for format_ext in PAINT3D_EXTS:
            try:
                shell_key_path = (f"SOFTWARE\\Classes\\SystemFileAssociations"
                                  f"\\.{format_ext}\\Shell")
                reg_utils.reg_create_key(shell_key_path, module=MOD)
                paint3d_key = f"{shell_key_path}\\3D Edit"
                reg_utils.reg_create_key(paint3d_key, module=MOD)
                reg_utils.reg_add(paint3d_key, '',
                                  "@%SystemRoot%\\system32\\mspaint.exe,-59500",
                                  'REG_SZ', module=MOD)
                command_key = f"{paint3d_key}\\command"
                reg_utils.reg_create_key(command_key, module=MOD)
                reg_utils.reg_add(
                    command_key, '',
                    '%SystemRoot%\\system32\\mspaint.exe "%1" /ForceBootstrapPaint3D',
                    'REG_EXPAND_SZ', module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR add_paint3d_edit: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_paint3d_edit():
    try:
        for format_ext in PAINT3D_EXTS:
            try:
                paint3d_key = (f"SOFTWARE\\Classes\\SystemFileAssociations"
                               f"\\.{format_ext}\\Shell\\3D Edit")
                reg_utils.reg_delete(paint3d_key, module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR remove_paint3d_edit: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")