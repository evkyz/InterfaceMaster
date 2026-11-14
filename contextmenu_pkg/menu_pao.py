"""Пункты, скрываемые через ProgrammaticAccessOnly (PAO).

Сюда входят: AppX-пакеты (Фотографии), wallpaper, print, pintohome,
powershell, а также пункты, скрываемые через Shell Extensions\\Blocked
(share_access, previous_version, compatibility, send_to_device,
modern_sharing) и через удаление ключей (copy_move, delete_folder_content,
take_ownership).
"""
import winreg
from tkinter import messagebox

import logger
from core import reg_utils
from contextmenu_pkg.constants import (
    WALLPAPER_EXTS, POWERSHELL_CLASSES, PINTOHOME_CLASSES,
    MODERN_SHARING_PATHS,
)
from contextmenu_pkg.menu_checks import (
    _iter_default_print_bases, _find_extra_print_bases,
    _appx_hkcu_path, _wallpaper_hkcu_path, _print_hkcu_path,
    _powershell_hkcu_path, _pintohome_hkcu_path,
)

MOD = "contextmenu"


# ---- PAO primitives ----

def _hide_key_pao(hkcu_path):
    try:
        reg_utils.reg_add('HKCU\\' + hkcu_path, 'ProgrammaticAccessOnly', '',
                          'REG_SZ', module=MOD)
        return True
    except Exception as e:
        logger.log(f"ERROR _hide_key_pao {hkcu_path}: {e}", module=MOD)
        return False


def _unhide_key_pao(hkcu_path):
    try:
        reg_utils.reg_delete('HKCU\\' + hkcu_path, 'ProgrammaticAccessOnly',
                             module=MOD)
        return True
    except Exception:
        return False


# ---- AppX (Photos) ----

def _hide_appx_for_user(kind):
    path = _appx_hkcu_path(kind)
    return _hide_key_pao(path) if path else False


def _unhide_appx_for_user(kind):
    path = _appx_hkcu_path(kind)
    return _unhide_key_pao(path) if path else False


def add_photos_edit():
    try:
        _unhide_appx_for_user("photos_edit")
    except Exception as e:
        logger.log(f"ERROR add_photos_edit: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_photos_edit():
    try:
        _hide_appx_for_user("photos_edit")
    except Exception as e:
        logger.log(f"ERROR remove_photos_edit: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


def add_create_video():
    try:
        _unhide_appx_for_user("create_video")
    except Exception as e:
        logger.log(f"ERROR add_create_video: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_create_video():
    try:
        _hide_appx_for_user("create_video")
    except Exception as e:
        logger.log(f"ERROR remove_create_video: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- Wallpaper ----

def _hide_wallpaper_for_user(ext):
    return _hide_key_pao(_wallpaper_hkcu_path(ext))


def _unhide_wallpaper_for_user(ext):
    return _unhide_key_pao(_wallpaper_hkcu_path(ext))


def add_wallpaper():
    try:
        for ext in WALLPAPER_EXTS:
            _unhide_wallpaper_for_user(ext)
    except Exception as e:
        logger.log(f"ERROR add_wallpaper: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_wallpaper():
    try:
        ok = 0
        for ext in WALLPAPER_EXTS:
            if _hide_wallpaper_for_user(ext):
                ok += 1
        if ok == 0:
            messagebox.showerror(
                "Ошибка",
                "Не удалось скрыть пункт.\n"
                "Попробуйте запустить программу от имени администратора.")
    except Exception as e:
        logger.log(f"ERROR remove_wallpaper: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- Print ----

def _hide_print_for_user(base):
    return _hide_key_pao(_print_hkcu_path(base))


def _unhide_print_for_user(base):
    return _unhide_key_pao(_print_hkcu_path(base))


def add_print():
    try:
        for base in _iter_default_print_bases():
            _unhide_print_for_user(base)
        for base in _find_extra_print_bases():
            _unhide_print_for_user(base)
    except Exception as e:
        logger.log(f"ERROR add_print: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_print():
    try:
        ok = 0
        for base in _iter_default_print_bases():
            if _hide_print_for_user(base):
                ok += 1
        for base in _find_extra_print_bases():
            if _hide_print_for_user(base):
                ok += 1
        if ok == 0:
            messagebox.showerror(
                "Ошибка",
                "Не удалось скрыть пункт.\n"
                "Попробуйте запустить программу от имени администратора.")
    except Exception as e:
        logger.log(f"ERROR remove_print: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- pintohome ----

def _hide_pintohome_for_user(base):
    return _hide_key_pao(_pintohome_hkcu_path(base))


def _unhide_pintohome_for_user(base):
    return _unhide_key_pao(_pintohome_hkcu_path(base))


def add_pintohome():
    try:
        for base in PINTOHOME_CLASSES:
            _unhide_pintohome_for_user(base)
    except Exception as e:
        logger.log(f"ERROR add_pintohome: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_pintohome():
    try:
        ok = 0
        for base in PINTOHOME_CLASSES:
            if _hide_pintohome_for_user(base):
                ok += 1
        if ok == 0:
            messagebox.showerror(
                "Ошибка",
                "Не удалось скрыть пункт.\n"
                "Попробуйте запустить программу от имени администратора.")
    except Exception as e:
        logger.log(f"ERROR remove_pintohome: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- PowerShell ----

def _hide_powershell_for_user(base):
    return _hide_key_pao(_powershell_hkcu_path(base))


def _unhide_powershell_for_user(base):
    return _unhide_key_pao(_powershell_hkcu_path(base))


def add_powershell():
    try:
        for base in POWERSHELL_CLASSES:
            _unhide_powershell_for_user(base)
    except Exception as e:
        logger.log(f"ERROR add_powershell: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_powershell():
    try:
        ok = 0
        for base in POWERSHELL_CLASSES:
            if _hide_powershell_for_user(base):
                ok += 1
        if ok == 0:
            messagebox.showerror(
                "Ошибка",
                "Не удалось скрыть пункт.\n"
                "Попробуйте запустить программу от имени администратора.")
    except Exception as e:
        logger.log(f"ERROR remove_powershell: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- Shell Extensions\Blocked (share_access / previous_version /
#      compatibility / send_to_device) ----

def add_share_access():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}")
                    logger.log_reg(MOD, "delete", blocked_path,
                                   "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}")
                except Exception:
                    pass
        except Exception:
            pass
    except Exception as e:
        logger.log(f"ERROR add_share_access: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении меню: {e}")


def remove_share_access():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        reg_utils.reg_add(blocked_path, "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}",
                          "", 'REG_SZ', module=MOD)
    except Exception as e:
        logger.log(f"ERROR remove_share_access: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении меню: {e}")


def add_previous_version():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{596AB062-B4D2-4215-9F74-E9109B0A8153}")
                    logger.log_reg(MOD, "delete", blocked_path,
                                   "{596AB062-B4D2-4215-9F74-E9109B0A8153}")
                except Exception:
                    pass
        except Exception:
            pass
    except Exception as e:
        logger.log(f"ERROR add_previous_version: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_previous_version():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        reg_utils.reg_add(blocked_path, "{596AB062-B4D2-4215-9F74-E9109B0A8153}",
                          "Восстановить прежнюю версию", 'REG_SZ', module=MOD)
    except Exception as e:
        logger.log(f"ERROR remove_previous_version: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


def add_compatibility():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{1d27f844-3a1f-4410-85ac-14651078412d}")
                    logger.log_reg(MOD, "delete", blocked_path,
                                   "{1d27f844-3a1f-4410-85ac-14651078412d}")
                except Exception:
                    pass
        except Exception:
            pass
    except Exception as e:
        logger.log(f"ERROR add_compatibility: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_compatibility():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        reg_utils.reg_add(blocked_path, "{1d27f844-3a1f-4410-85ac-14651078412d}",
                          "Исправление проблем с совместимостью", 'REG_SZ',
                          module=MOD)
    except Exception as e:
        logger.log(f"ERROR remove_compatibility: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


def add_send_to_device():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, blocked_path, 0,
                                winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}")
                    logger.log_reg(MOD, "delete", blocked_path,
                                   "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}")
                except FileNotFoundError:
                    pass
        except FileNotFoundError:
            pass
    except Exception as e:
        logger.log(f"ERROR add_send_to_device: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_send_to_device():
    try:
        blocked_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
        reg_utils.reg_add(blocked_path, "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}",
                          "Передать на устройство", 'REG_SZ', module=MOD)
    except Exception as e:
        logger.log(f"ERROR remove_send_to_device: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


# ---- modern_sharing / copy_move / delete_folder_content / take_ownership ----

def add_modern_sharing():
    try:
        clsid = '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'
        for key_path in MODERN_SHARING_PATHS:
            try:
                reg_utils.reg_add(key_path, '', clsid, 'REG_SZ', module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR add_modern_sharing: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_modern_sharing():
    try:
        for key_path in MODERN_SHARING_PATHS:
            try:
                reg_utils.reg_delete(key_path, module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR remove_modern_sharing: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


def add_copy_move_menu():
    try:
        for key_path, clsid in (
            (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
             "{C2FBB630-2971-11D1-A18C-00C04FD75D13}"),
            (r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To',
             "{C2FBB631-2971-11D1-A18C-00C04FD75D13}"),
        ):
            try:
                reg_utils.reg_add(key_path, '', clsid, 'REG_SZ', module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR add_copy_move_menu: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении меню: {e}")


def remove_copy_move_menu():
    try:
        for key_path in (
            r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To',
            r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Move To',
        ):
            try:
                reg_utils.reg_delete(key_path, module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR remove_copy_move_menu: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении меню: {e}")


def add_delete_folder_content():
    try:
        key_path = r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'
        reg_utils.reg_add(key_path, '', 'Удалить содержимое папки',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path, 'Icon', 'shell32.dll,-240',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(key_path + r'\command', '',
                          'cmd.exe /c "cd /d \\"%1\\" && del /q /f /s *.*"',
                          'REG_SZ', module=MOD)
    except Exception as e:
        logger.log(f"ERROR add_delete_folder_content: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_delete_folder_content():
    try:
        reg_utils.reg_delete(
            r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent', module=MOD)
    except Exception as e:
        logger.log(f"ERROR remove_delete_folder_content: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")


def add_take_ownership():
    try:
        for base_key, cmd in (
            (r'*\shell\runas',
             'cmd.exe /c takeown /f "%1" && icacls "%1" /grant администраторы:F'),
            (r'Directory\shell\runas',
             'cmd.exe /c takeown /f "%1" /r /d y && icacls "%1" /grant администраторы:F /t'),
        ):
            reg_utils.reg_add(base_key, '', 'Стать владельцем',
                              'REG_SZ', module=MOD)
            reg_utils.reg_add(base_key, 'NoWorkingDirectory', '',
                              'REG_SZ', module=MOD)
            reg_utils.reg_add(base_key, 'Icon', 'imageres.dll,54',
                              'REG_SZ', module=MOD)
            reg_utils.reg_add(base_key, 'MUIVerb', 'Стать владельцем',
                              'REG_SZ', module=MOD)
            reg_utils.reg_add(base_key + r'\command', '', cmd,
                              'REG_SZ', module=MOD)
            reg_utils.reg_add(base_key + r'\command', 'IsolatedCommand', cmd,
                              'REG_SZ', module=MOD)
    except Exception as e:
        logger.log(f"ERROR add_take_ownership: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при добавлении пункта: {e}")


def remove_take_ownership():
    try:
        for key_path in (r'*\shell\runas', r'Directory\shell\runas'):
            try:
                reg_utils.reg_delete(key_path, module=MOD)
            except Exception:
                pass
    except Exception as e:
        logger.log(f"ERROR remove_take_ownership: {e}", module=MOD)
        messagebox.showerror("Ошибка", f"Ошибка при удалении пункта: {e}")