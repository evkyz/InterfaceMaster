"""Полный сброс настроек к стандартным значениям Windows."""
import hashlib
import os
import subprocess
import time
import tkinter as tk
import winreg
from tkinter import ttk, messagebox

import logger
from core import reg_utils, win_utils

MOD = "settings_reset"

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

DRIVE_ICONS = r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons"
RECYCLE_PATH = r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon"

MOD_EXPECTED_HASHES = {
    "imageres.dll": "DDD5D7462A26F7837CD02A7667615354F9377A0A727BF880D825676D6DDC4CB7",
    "imagesp1.dll": "466E64D491C7B3D7051FBDC0907CD813D7A6435D18BF4B00CE0C3191B80E0541",
    "networkexplorer.dll": "7BF68211CFBDDE2E377C992CEC2F5A0C1F41B1F12BE3A99733A549AB63A8C961",
    "DDORes.dll": "4802C8F0BAAE3765FDFF16F09F56647B1C8B20889E4C531B773442EC068CBC3B",
    "mmres.dll": "4651BB26196D0ED493A26F4E401C8FAC0465E7F9DE2ECFED1642155A6EAF3703",
}


def check_modified_files():
    modified = []
    for fname, expected in MOD_EXPECTED_HASHES.items():
        path = os.path.join(os.environ["WINDIR"], "System32", fname)
        if not os.path.exists(path):
            continue
        try:
            with open(path, "rb") as f:
                if hashlib.sha256(f.read()).hexdigest().upper() == expected.upper():
                    modified.append(fname)
        except Exception:
            pass
    return modified


def _restore_modified_files():
    modified_names = []
    for fname in MOD_EXPECTED_HASHES:
        path = os.path.join(os.environ["WINDIR"], "System32", fname)
        if not os.path.exists(path):
            continue
        try:
            with open(path, "rb") as f:
                h = hashlib.sha256(f.read()).hexdigest().upper()
            if h == MOD_EXPECTED_HASHES[fname].upper():
                modified_names.append(fname)
        except Exception:
            pass

    if not modified_names:
        return

    try:
        subprocess.run("taskkill /f /im explorer.exe", shell=True,
                       capture_output=True, creationflags=CREATE_NO_WINDOW)
        time.sleep(0.5)
    except Exception:
        pass

    for fname in modified_names:
        path = os.path.join(os.environ["WINDIR"], "System32", fname)
        backup = f"{path}_orig"
        try:
            subprocess.run(f'takeown /f "{path}" /a >nul 2>&1',
                           shell=True, capture_output=True,
                           creationflags=CREATE_NO_WINDOW)
            subprocess.run(f'icacls "{path}" /grant "%USERNAME%":(F) >nul 2>&1',
                           shell=True, capture_output=True,
                           creationflags=CREATE_NO_WINDOW)
            os.remove(path)
            if os.path.exists(backup):
                os.rename(backup, path)
                subprocess.run(
                    f'icacls "{path}" /setowner "NT SERVICE\\TrustedInstaller" >nul 2>&1',
                    shell=True, capture_output=True,
                    creationflags=CREATE_NO_WINDOW)
            logger.log(f"restored {fname}", module=MOD)
        except Exception as e:
            logger.log(f"ERROR restore {fname}: {e}", module=MOD)

    try:
        subprocess.Popen(['explorer.exe'], creationflags=CREATE_NO_WINDOW)
    except Exception:
        pass


def _reset_desktop_icons_direct():
    defaults = (
        (r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID"
         r"\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon",
         r'C:\Windows\System32\imageres.dll,-123'),
        (r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID"
         r"\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon",
         r'C:\Windows\System32\imageres.dll,-109'),
        (r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID"
         r"\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon",
         r'C:\Windows\System32\imageres.dll,-25'),
    )
    for path, icon in defaults:
        reg_utils.reg_add('HKCU\\' + path, None, icon, 'REG_SZ', module=MOD)

    reg_utils.reg_add('HKCU\\' + RECYCLE_PATH, None,
                      r'C:\Windows\System32\imageres.dll,-55',
                      'REG_SZ', module=MOD)
    reg_utils.reg_add('HKCU\\' + RECYCLE_PATH, 'Empty',
                      r'C:\Windows\System32\imageres.dll,-55',
                      'REG_SZ', module=MOD)
    reg_utils.reg_add('HKCU\\' + RECYCLE_PATH, 'Full',
                      r'C:\Windows\System32\imageres.dll,-54',
                      'REG_SZ', module=MOD)

    cp_defaults = {
        '{26EE0668-A00A-44D7-9371-BEB064C98683}': r'C:\Windows\System32\imageres.dll,-21',
        '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}': r'C:\Windows\System32\imageres.dll,-27',
        '{21EC2020-3AEA-1069-A2DD-08002B30309D}': r'C:\Windows\System32\imageres.dll,-21',
    }
    for clsid, icon in cp_defaults.items():
        path = (r"Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID"
                rf"\{clsid}\DefaultIcon")
        reg_utils.reg_add('HKCU\\' + path, None, icon, 'REG_SZ', module=MOD)

    lib_value = r'C:\Windows\System32\imageres.dll,-1023'
    for hive, path in (
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
        (winreg.HKEY_CURRENT_USER,
         r"Software\Classes\CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
        (winreg.HKEY_CLASSES_ROOT,
         r"CLSID\{031E4825-7B94-4dc3-B131-E946B44C8DD5}\DefaultIcon"),
    ):
        try:
            try:
                k = winreg.OpenKey(hive, path, 0, winreg.KEY_WRITE)
            except FileNotFoundError:
                k = winreg.CreateKey(hive, path)
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, lib_value)
            winreg.CloseKey(k)
            break
        except (PermissionError, OSError):
            continue
        except Exception:
            continue

    subprocess.run(['ie4uinit.exe', '-show'], capture_output=True,
                   creationflags=CREATE_NO_WINDOW)


def reset_all_settings(parent):
    modified = check_modified_files()
    msg = ("ВНИМАНИЕ! Это действие сбросит все настройки к стандартным значениям Windows.\n\n"
           "Будут сброшены:\n• Иконки рабочего стола\n• Иконки дисков\n"
           "• Стрелки ярлыков\n• Текст '- ярлык'\n• Пункты контекстного меню\n"
           "• Настройки Проводника\n• Настройки Панели задач\n"
           "• Пункты меню 'Создать'\n• Пункты меню 'Отправить'\n")
    if modified:
        msg += "\n⚠️ БУДУТ ВОССТАНОВЛЕНЫ МОДИФИЦИРОВАННЫЕ ФАЙЛЫ:\n"
        for f in modified:
            msg += f"  • {f}\n"
    msg += "\n\nПродолжить?"

    if not messagebox.askyesno("Сброс всех настроек", msg, icon='warning'):
        return

    logger.log("reset all settings", module=MOD)
    errors = []

    if modified:
        try:
            _restore_modified_files()
        except Exception as e:
            errors.append(f"restore dll: {e}")

    try:
        _reset_desktop_icons_direct()
    except Exception as e:
        errors.append(f"desktop: {e}")

    try:
        reg_utils.reg_delete(DRIVE_ICONS, module=MOD)
    except Exception as e:
        errors.append(f"disk: {e}")

    # str
    try:
        from str import Module as StrModule
        temp = ttk.Frame(parent)
        s = StrModule(temp)
        temp.pack_forget()
        s.arrow_choice.set(1)
        s.shortcut_text_choice.set(1)
        s.apply_changes()
        s.hide()
        temp.destroy()
    except Exception as e:
        errors.append(f"str: {e}")

    # contextmenu
    try:
        from contextmenu import DesktopMenuTab
        temp = ttk.Frame(parent)
        dt = DesktopMenuTab(temp)
        dt.set_default_settings()
        temp.destroy()
    except Exception as e:
        errors.append(f"contextmenu: {e}")

    # explorer
    try:
        from explorer import ExplorerModule, IMAGE_DETAILS_KEY, IMAGE_DETAILS_VALUE
        temp = ttk.Frame(parent)
        e = ExplorerModule(temp)
        temp.pack_forget()
        is_admin = win_utils.is_admin()

        if is_admin:
            for guid in e.checkbox_vars:
                e.checkbox_vars[guid].set(True)
            e.apply_folder_changes()

        if is_admin and not e.bitlocker_enabled:
            e.toggle_bitlocker()
        if is_admin and e.photo_viewer_enabled:
            e.toggle_photo_viewer()
        if e.file_ext_enabled:
            e.toggle_file_ext()
        if not e.recent_folders_enabled:
            e.toggle_recent_folders()

        if getattr(e, 'image_details_supported', False):
            if getattr(e, 'image_details_hidden', False):
                e.toggle_image_details()
        else:
            reg_utils.reg_delete(IMAGE_DETAILS_KEY, IMAGE_DETAILS_VALUE, module=MOD)

        e.hide()
        temp.destroy()
    except Exception as e:
        errors.append(f"explorer: {e}")

    # taskbar
    try:
        from taskbar import Module as TaskbarModule
        temp = ttk.Frame(parent)
        t = TaskbarModule(temp)
        temp.pack_forget()
        t._reset_all()
        t.hide()
        temp.destroy()
    except Exception as e:
        errors.append(f"taskbar: {e}")

    # createnew
    try:
        from contextmenu_pkg.createnew import (
            _reg_exists, _rename_key,
            TARGET_EXTENSIONS, MS_EXTENSIONS,
            _ext_registered, _find_shellnew_under_ext,
        )
        import winreg as _wr

        def _enable(hive, base):
            try:
                if _reg_exists(hive, base + '_no'):
                    logger.log_reg(MOD, "rename", base + '_no', None, base)
                    _rename_key(hive, base + '_no', base)
            except Exception:
                pass

        _enable(_wr.HKEY_CLASSES_ROOT, r'Folder\ShellNew')
        for _, path in TARGET_EXTENSIONS.items():
            _enable(_wr.HKEY_LOCAL_MACHINE, path)

        for ext in MS_EXTENSIONS:
            if not _ext_registered(ext):
                continue
            for hive in (_wr.HKEY_LOCAL_MACHINE, _wr.HKEY_CURRENT_USER):
                for h, path, cls, enabled in _find_shellnew_under_ext(hive, ext):
                    if path.endswith('_no'):
                        try:
                            logger.log_reg(MOD, "rename", path, None, path[:-3])
                            _rename_key(h, path, path[:-3])
                        except Exception:
                            pass
    except Exception as e:
        errors.append(f"createnew: {e}")

    # sendto
    try:
        import ctypes
        sendto_path = os.path.join(os.environ.get('APPDATA', ''),
                                   'Microsoft', 'Windows', 'SendTo')
        if os.path.isdir(sendto_path):
            for name in os.listdir(sendto_path):
                if name.lower() == 'desktop.ini':
                    continue
                full = os.path.join(sendto_path, name)
                if not (os.path.isfile(full) or os.path.islink(full)):
                    continue
                attrs = ctypes.windll.kernel32.GetFileAttributesW(full)
                if attrs == -1:
                    continue
                ctypes.windll.kernel32.SetFileAttributesW(full, attrs & ~0x2)
                logger.log_file_attr(MOD, full, "unhide")
    except Exception as e:
        errors.append(f"sendto: {e}")

    win_utils.restart_explorer(MOD)

    # ShellState
    try:
        from explorer import (
            SHELLSTATE_KEY, SHELLSTATE_VALUE, SHELLSTATE_CONFIRM_OFF,
        )
        time.sleep(0.3)
        default_data = bytearray((
            0x24, 0x00, 0x00, 0x00,
            0x00, 0x28, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x00, 0x00,
            0x13, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x62, 0x00, 0x00, 0x00,
        ))
        default_data[4] = SHELLSTATE_CONFIRM_OFF
        logger.log_reg(MOD, "set", SHELLSTATE_KEY, SHELLSTATE_VALUE, "byte4=0x37")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, SHELLSTATE_KEY) as key:
            winreg.SetValueEx(key, SHELLSTATE_VALUE, 0,
                              winreg.REG_BINARY, bytes(default_data))
        time.sleep(0.3)
        win_utils.restart_explorer(MOD)
    except Exception as e:
        errors.append(f"delete_confirm final: {e}")

    if errors:
        messagebox.showwarning("Сброс завершён с ошибками",
                               "\n".join(errors[:5]))