"""Проверки существования/скрытости пунктов контекстного меню."""
import winreg

from contextmenu_pkg.constants import (
    MODERN_SHARING_PATHS, WALLPAPER_EXTS, POWERSHELL_CLASSES,
    PINTOHOME_CLASSES, PHOTOS_SHELL_EDIT, PHOTOS_SHELL_CREATE_VIDEO,
    DEFAULT_PRINT_BASES,
)

_PRINT_EXTRA_BASES_CACHE = None


# ---- AppX ----

def _appx_hkcu_path(kind):
    if kind == "photos_edit":
        return f"Software\\Classes\\{PHOTOS_SHELL_EDIT}"
    if kind == "create_video":
        return f"Software\\Classes\\{PHOTOS_SHELL_CREATE_VIDEO}"
    return None


def _appx_is_hidden(kind):
    path = _appx_hkcu_path(kind)
    if not path:
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            try:
                winreg.QueryValueEx(key, "ProgrammaticAccessOnly")
                return True
            except FileNotFoundError:
                return False
    except (FileNotFoundError, OSError):
        return False


# ---- print ----

def _iter_default_print_bases():
    result = []
    for base in DEFAULT_PRINT_BASES:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{base}\\shell\\print"):
                result.append(base)
        except (FileNotFoundError, OSError):
            pass
    return result


def _scan_extra_print_bases():
    default_lower = {b.lower() for b in DEFAULT_PRINT_BASES}
    extra = []
    seen = set()

    def _add(base):
        if not base:
            return
        low = base.lower()
        if low in default_lower or low in seen:
            return
        seen.add(low)
        extra.append(base)

    def _scan_key(path):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, path) as key:
                try:
                    with winreg.OpenKey(key, "shell") as shell_key:
                        i = 0
                        while True:
                            try:
                                sub = winreg.EnumKey(shell_key, i)
                            except OSError:
                                break
                            i += 1
                            if sub.lower() == "print":
                                return True
                except (FileNotFoundError, OSError):
                    pass
        except (FileNotFoundError, OSError):
            pass
        return False

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as root:
            i = 0
            while True:
                try:
                    top = winreg.EnumKey(root, i)
                except OSError:
                    break
                i += 1
                low = top.lower()
                if low.startswith((".", "clsid", "interface", "typelib",
                                   "wow6432node", "applications")):
                    continue
                if _scan_key(top):
                    _add(top)
    except (FileNotFoundError, OSError):
        pass

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "SystemFileAssociations") as sfa:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(sfa, i)
                except OSError:
                    break
                i += 1
                base = f"SystemFileAssociations\\{sub}"
                if _scan_key(base):
                    _add(base)
    except (FileNotFoundError, OSError):
        pass

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Applications") as apps:
            i = 0
            while True:
                try:
                    exe = winreg.EnumKey(apps, i)
                except OSError:
                    break
                i += 1
                base = f"Applications\\{exe}"
                if _scan_key(base):
                    _add(base)
    except (FileNotFoundError, OSError):
        pass

    return extra


def _find_extra_print_bases():
    global _PRINT_EXTRA_BASES_CACHE
    if _PRINT_EXTRA_BASES_CACHE is not None:
        return _PRINT_EXTRA_BASES_CACHE
    _PRINT_EXTRA_BASES_CACHE = _scan_extra_print_bases()
    return _PRINT_EXTRA_BASES_CACHE


def invalidate_print_cache():
    global _PRINT_EXTRA_BASES_CACHE
    _PRINT_EXTRA_BASES_CACHE = None


# ---- PAO checkers ----

def _wallpaper_hkcu_path(ext):
    return f"Software\\Classes\\SystemFileAssociations\\.{ext}\\Shell\\setdesktopwallpaper"


def _wallpaper_is_hidden(ext):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _wallpaper_hkcu_path(ext)) as key:
            try:
                winreg.QueryValueEx(key, "ProgrammaticAccessOnly")
                return True
            except FileNotFoundError:
                return False
    except (FileNotFoundError, OSError):
        return False


def _print_hkcu_path(base):
    return f"Software\\Classes\\{base}\\shell\\print"


def _print_is_hidden(base):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _print_hkcu_path(base)) as key:
            try:
                winreg.QueryValueEx(key, "ProgrammaticAccessOnly")
                return True
            except FileNotFoundError:
                return False
    except (FileNotFoundError, OSError):
        return False


def _powershell_hkcu_path(base):
    return f"Software\\Classes\\{base}\\shell\\PowerShell"


def _powershell_is_hidden(base):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _powershell_hkcu_path(base)) as key:
            try:
                winreg.QueryValueEx(key, "ProgrammaticAccessOnly")
                return True
            except FileNotFoundError:
                return False
    except (FileNotFoundError, OSError):
        return False


def _pintohome_hkcu_path(base):
    return f"Software\\Classes\\{base}\\shell\\pintohome"


def _pintohome_is_hidden(base):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _pintohome_hkcu_path(base)) as key:
            try:
                winreg.QueryValueEx(key, "ProgrammaticAccessOnly")
                return True
            except FileNotFoundError:
                return False
    except (FileNotFoundError, OSError):
        return False


# ---- supported checks (прокси в core.win_utils) ----

def is_paint3d_supported():
    from core import win_utils
    return win_utils.is_paint3d_supported()


def is_photos_supported():
    from core import win_utils
    return win_utils.is_photos_supported()


def is_create_video_supported():
    from core import win_utils
    return win_utils.is_create_video_supported()


def is_ltsc_edition():
    from core import win_utils
    return win_utils.is_ltsc_edition()


# ---- главная проверка ----

def check_menu_item_exists(item_type):
    try:
        if item_type == "admin_menu":
            for hive, path in (
                (winreg.HKEY_CURRENT_USER, r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'),
            ):
                try:
                    with winreg.OpenKey(hive, path) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Администрирование"
                except FileNotFoundError:
                    continue
            return False

        if item_type == "system_menu":
            for hive, path in (
                (winreg.HKEY_CURRENT_USER, r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'),
                (winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'),
            ):
                try:
                    with winreg.OpenKey(hive, path) as key:
                        value, _ = winreg.QueryValueEx(key, "MUIVerb")
                        return value == "Система"
                except FileNotFoundError:
                    continue
            return False

        if item_type == "regedit":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'DesktopBackground\Shell\Regedit') as key:
                    value, _ = winreg.QueryValueEx(key, "MUIVerb")
                    return value == "Редактор реестра"
            except FileNotFoundError:
                return False

        if item_type == "restart_explorer":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'DesktopBackground\Shell\RestartExplorer') as key:
                    value, _ = winreg.QueryValueEx(key, "MUIVerb")
                    return value == "Перезапустить explorer"
            except FileNotFoundError:
                return False

        if item_type == "clipboard_clear":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'DesktopBackground\Shell\Clip') as key:
                    value, _ = winreg.QueryValueEx(key, "MUIVerb")
                    return value == "Очистить буфер"
            except FileNotFoundError:
                return False

        if item_type == "power_menu":
            for key_path in (
                r'DesktopBackground\Shell\Hibernation',
                r'DesktopBackground\Shell\Reboot',
                r'DesktopBackground\Shell\Shutdown',
            ):
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                        return True
                except Exception:
                    continue
            return False

        if item_type == "task_manager":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'DesktopBackground\Shell\TaskManager') as key:
                    value, _ = winreg.QueryValueEx(key, "MUIVerb")
                    return value == "Диспетчер задач"
            except FileNotFoundError:
                return False

        if item_type == "copy_as_path":
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu') as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    return value == "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}"
            except FileNotFoundError:
                return False

        if item_type == "pin_to_start":
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r'SOFTWARE\Classes\exefile\shellex\ContextMenuHandlers\PintoStartScreen') as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    return value == "{470C0EBD-5D73-4d58-9CED-E91E22E23282}"
            except FileNotFoundError:
                return False

        if item_type == "pin_to_taskbar":
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\{90AA3A4E-1CBA-4233-B8BB-535773D48449}') as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    return value == "Taskband Pin"
            except FileNotFoundError:
                return False

        if item_type in ("send_to_device", "compatibility", "share_access", "previous_version"):
            guid = {
                "send_to_device": "{7AD84985-87B4-4a16-BE58-8B72A5B390F7}",
                "compatibility": "{1d27f844-3a1f-4410-85ac-14651078412d}",
                "share_access": "{f81e9010-6ea4-11ce-a7ff-00aa003ca9f6}",
                "previous_version": "{596AB062-B4D2-4215-9F74-E9109B0A8153}",
            }[item_type]
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked') as key:
                    try:
                        winreg.QueryValueEx(key, guid)
                        return False
                    except FileNotFoundError:
                        return True
            except FileNotFoundError:
                return True
            except Exception:
                return True

        if item_type == "paint3d":
            for format_ext in ('bmp', 'fbx', 'gif'):
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                        f"SOFTWARE\\Classes\\SystemFileAssociations\\.{format_ext}\\Shell\\3D Edit"):
                        return True
                except Exception:
                    continue
            return False

        if item_type == "photos_edit":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, PHOTOS_SHELL_EDIT):
                    pass
            except (FileNotFoundError, OSError):
                return False
            return not _appx_is_hidden("photos_edit")

        if item_type == "create_video":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, PHOTOS_SHELL_CREATE_VIDEO):
                    pass
            except (FileNotFoundError, OSError):
                return False
            return not _appx_is_hidden("create_video")

        if item_type == "wallpaper":
            for ext in WALLPAPER_EXTS:
                try:
                    with winreg.OpenKey(
                            winreg.HKEY_CLASSES_ROOT,
                            f"SystemFileAssociations\\.{ext}\\Shell\\setdesktopwallpaper"):
                        pass
                except (FileNotFoundError, OSError):
                    continue
                if not _wallpaper_is_hidden(ext):
                    return True
            return False

        if item_type == "print":
            for base in _iter_default_print_bases():
                if not _print_is_hidden(base):
                    return True
            for base in _find_extra_print_bases():
                if not _print_is_hidden(base):
                    return True
            return False

        if item_type == "pintohome":
            for base in PINTOHOME_CLASSES:
                try:
                    with winreg.OpenKey(
                            winreg.HKEY_CLASSES_ROOT,
                            f"{base}\\shell\\pintohome"):
                        pass
                except (FileNotFoundError, OSError):
                    continue
                if not _pintohome_is_hidden(base):
                    return True
            return False

        if item_type == "powershell":
            for base in POWERSHELL_CLASSES:
                try:
                    with winreg.OpenKey(
                            winreg.HKEY_CLASSES_ROOT,
                            f"{base}\\shell\\PowerShell"):
                        pass
                except (FileNotFoundError, OSError):
                    continue
                if not _powershell_is_hidden(base):
                    return True
            return False

        if item_type == "defender":
            defender_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
            defender_no = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no'
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_path):
                    return True
            except FileNotFoundError:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_no):
                        return False
                except FileNotFoundError:
                    return False
            except Exception:
                return False

        if item_type == "modern_sharing":
            for path in MODERN_SHARING_PATHS:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                        value, _ = winreg.QueryValueEx(key, "")
                        if value == "{e2bf9676-5f8f-435c-97eb-11607a5bedf7}":
                            return True
                except Exception:
                    continue
            return False

        if item_type == "delete_folder_content":
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r'SOFTWARE\Classes\Directory\shell\DeleteFolderContent'):
                    return True
            except FileNotFoundError:
                return False

        if item_type == "copy_move":
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\Copy To'):
                    return True
            except FileNotFoundError:
                return False

        if item_type == "take_ownership":
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'*\shell\runas'):
                    return True
            except FileNotFoundError:
                return False
    except Exception:
        return False
    return False