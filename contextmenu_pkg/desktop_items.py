"""Определения пунктов меню рабочего стола (только данные)."""
from contextmenu_pkg import (
    menu_pao, menu_admin, menu_system, menu_power, menu_registry_ops,
)
from contextmenu_pkg.menu_checks import (
    check_menu_item_exists, is_paint3d_supported, is_photos_supported,
    is_create_video_supported,
)
from contextmenu_pkg.constants import TIP_DESKTOP, TIP_COMPUTER


def build_menu_items():
    return [
        {"name": "Закрепить на панели быстрого доступа",
         "add_command": menu_pao.add_pintohome,
         "remove_command": menu_pao.remove_pintohome,
         "check_command": lambda: check_menu_item_exists("pintohome"),
         "enabled": False, "default_enabled": True},
        {"name": "Открыть окно PowerShell здесь",
         "add_command": menu_pao.add_powershell,
         "remove_command": menu_pao.remove_powershell,
         "check_command": lambda: check_menu_item_exists("powershell"),
         "enabled": False, "default_enabled": True},
        {"name": "Предоставить доступ к",
         "add_command": menu_pao.add_share_access,
         "remove_command": menu_pao.remove_share_access,
         "check_command": lambda: check_menu_item_exists("share_access"),
         "enabled": False, "default_enabled": True},
        {"name": "Отправить (metro)",
         "add_command": menu_pao.add_modern_sharing,
         "remove_command": menu_pao.remove_modern_sharing,
         "check_command": lambda: check_menu_item_exists("modern_sharing"),
         "enabled": False, "default_enabled": True},
        {"name": "Проверка с использованием Windows Defender",
         "add_command": _add_defender,
         "remove_command": _remove_defender,
         "check_command": lambda: check_menu_item_exists("defender"),
         "enabled": False, "default_enabled": True},
        {"name": "Восстановить прежнюю версию",
         "add_command": menu_pao.add_previous_version,
         "remove_command": menu_pao.remove_previous_version,
         "check_command": lambda: check_menu_item_exists("previous_version"),
         "enabled": False, "default_enabled": True},
        {"name": "Закрепить на начальном экране",
         "add_command": menu_registry_ops.add_pin_to_start,
         "remove_command": menu_registry_ops.remove_pin_to_start,
         "check_command": lambda: check_menu_item_exists("pin_to_start"),
         "enabled": False, "default_enabled": True},
        {"name": "Закрепить на Панели задач",
         "add_command": menu_registry_ops.add_pin_to_taskbar,
         "remove_command": menu_registry_ops.remove_pin_to_taskbar,
         "check_command": lambda: check_menu_item_exists("pin_to_taskbar"),
         "enabled": False, "default_enabled": True},
        {"name": "Копировать как путь",
         "add_command": menu_registry_ops.add_copy_as_path,
         "remove_command": menu_registry_ops.remove_copy_as_path,
         "check_command": lambda: check_menu_item_exists("copy_as_path"),
         "enabled": False, "default_enabled": True},
        {"name": "Исправление проблем с совместимостью",
         "add_command": menu_pao.add_compatibility,
         "remove_command": menu_pao.remove_compatibility,
         "check_command": lambda: check_menu_item_exists("compatibility"),
         "enabled": False, "default_enabled": True},
        {"name": "Изменить с помощью Paint 3D",
         "add_command": menu_registry_ops.add_paint3d_edit,
         "remove_command": menu_registry_ops.remove_paint3d_edit,
         "check_command": lambda: check_menu_item_exists("paint3d"),
         "enabled": False, "default_enabled": True,
         "supported_check": is_paint3d_supported,
         "unsupported_tip": "Доступно в Windows 10 1703 и выше"},
        {"name": "Изменить с помощью Фотографии",
         "add_command": menu_pao.add_photos_edit,
         "remove_command": menu_pao.remove_photos_edit,
         "check_command": lambda: check_menu_item_exists("photos_edit"),
         "enabled": False, "default_enabled": True,
         "supported_check": is_photos_supported,
         "unsupported_tip": "Доступно в Windows 10 1703 и выше"},
        {"name": "Создать новое видео",
         "add_command": menu_pao.add_create_video,
         "remove_command": menu_pao.remove_create_video,
         "check_command": lambda: check_menu_item_exists("create_video"),
         "enabled": False, "default_enabled": True,
         "supported_check": is_create_video_supported,
         "unsupported_tip": "Доступно в Windows 10 1809 и выше"},
        {"name": "Сделать фоном рабочего стола",
         "add_command": menu_pao.add_wallpaper,
         "remove_command": menu_pao.remove_wallpaper,
         "check_command": lambda: check_menu_item_exists("wallpaper"),
         "enabled": False, "default_enabled": True},
        {"name": "Печать",
         "add_command": menu_pao.add_print,
         "remove_command": menu_pao.remove_print,
         "check_command": lambda: check_menu_item_exists("print"),
         "enabled": False, "default_enabled": True},
        {"name": "Передать на устройство",
         "add_command": menu_pao.add_send_to_device,
         "remove_command": menu_pao.remove_send_to_device,
         "check_command": lambda: check_menu_item_exists("send_to_device"),
         "enabled": False, "default_enabled": True},
        {"name": "Стать владельцем",
         "add_command": menu_pao.add_take_ownership,
         "remove_command": menu_pao.remove_take_ownership,
         "check_command": lambda: check_menu_item_exists("take_ownership"),
         "enabled": False, "default_enabled": False},
        {"name": "Копировать/Переместить в папку",
         "add_command": menu_pao.add_copy_move_menu,
         "remove_command": menu_pao.remove_copy_move_menu,
         "check_command": lambda: check_menu_item_exists("copy_move"),
         "enabled": False, "default_enabled": False},
        {"name": "Удалить содержимое папки",
         "add_command": menu_pao.add_delete_folder_content,
         "remove_command": menu_pao.remove_delete_folder_content,
         "check_command": lambda: check_menu_item_exists("delete_folder_content"),
         "enabled": False, "default_enabled": False},
        {"name": "Диспетчер задач", "add_tooltip": TIP_DESKTOP,
         "add_command": menu_registry_ops.add_task_manager,
         "remove_command": menu_registry_ops.remove_task_manager,
         "check_command": lambda: check_menu_item_exists("task_manager"),
         "enabled": False, "default_enabled": False},
        {"name": "Очистить буфер", "add_tooltip": TIP_DESKTOP,
         "add_command": menu_registry_ops.add_clipboard_clear,
         "remove_command": menu_registry_ops.remove_clipboard_clear,
         "check_command": lambda: check_menu_item_exists("clipboard_clear"),
         "enabled": False, "default_enabled": False},
        {"name": "Перезапустить explorer", "add_tooltip": TIP_DESKTOP,
         "add_command": menu_registry_ops.add_restart_explorer,
         "remove_command": menu_registry_ops.remove_restart_explorer,
         "check_command": lambda: check_menu_item_exists("restart_explorer"),
         "enabled": False, "default_enabled": False},
        {"name": "Редактор реестра", "add_tooltip": TIP_DESKTOP,
         "add_command": menu_registry_ops.add_regedit,
         "remove_command": menu_registry_ops.remove_regedit,
         "check_command": lambda: check_menu_item_exists("regedit"),
         "enabled": False, "default_enabled": False},
        {"name": "Выключение/Перезагрузка/Гибернация", "add_tooltip": TIP_DESKTOP,
         "add_command": menu_power.add_power_menu,
         "remove_command": menu_power.remove_power_menu,
         "check_command": lambda: check_menu_item_exists("power_menu"),
         "enabled": False, "default_enabled": False},
        {"name": "Администрирование", "add_tooltip": TIP_COMPUTER,
         "add_command": menu_admin.add_admin_menu,
         "remove_command": menu_admin.remove_admin_menu,
         "check_command": lambda: check_menu_item_exists("admin_menu"),
         "enabled": False, "default_enabled": False},
        {"name": "Система", "add_tooltip": TIP_COMPUTER,
         "add_command": menu_system.add_system_menu,
         "remove_command": menu_system.remove_system_menu,
         "check_command": lambda: check_menu_item_exists("system_menu"),
         "enabled": False, "default_enabled": False},
    ]


# ---- defender: реализация (перенесена из menu_pao/constants) ----

def _add_defender():
    import winreg
    from tkinter import messagebox
    from core import reg_utils
    from contextmenu_pkg.menu_registry_helpers import rename_registry_key
    import logger
    try:
        defender_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
        defender_no = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no'
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_no):
                if rename_registry_key(winreg.HKEY_LOCAL_MACHINE, defender_no,
                                       defender_path):
                    logger.log_reg("contextmenu", "rename", defender_no, None,
                                   defender_path)
                    return
        except Exception:
            pass
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_path):
                return
        except Exception:
            pass
        try:
            reg_utils.reg_add(defender_path, '',
                              'Microsoft.Windows.Defender.IconHandler',
                              'REG_SZ', module="contextmenu")
            inproc_path = defender_path + r'\InprocServer32'
            reg_utils.reg_add(
                inproc_path, '',
                '%SystemRoot%\\System32\\WindowsDefender\\ShellExt.dll',
                'REG_EXPAND_SZ', module="contextmenu")
            reg_utils.reg_add(inproc_path, 'ThreadingModel', 'Apartment',
                              'REG_SZ', module="contextmenu")
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при добавлении пункта")
    except Exception:
        messagebox.showerror("Ошибка", "Ошибка при добавлении пункта")


def _remove_defender():
    import winreg
    from tkinter import messagebox
    from contextmenu_pkg.menu_registry_helpers import rename_registry_key
    import logger
    try:
        defender_path = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}'
        defender_no = r'SOFTWARE\Classes\CLSID\{09A47860-11B0-4DA5-AFA5-26D86198A780}_no'
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, defender_path):
                if rename_registry_key(winreg.HKEY_LOCAL_MACHINE, defender_path,
                                       defender_no):
                    logger.log_reg("contextmenu", "rename", defender_path, None,
                                   defender_no)
                    return True
                else:
                    messagebox.showerror("Ошибка", "Не удалось удалить пункт")
                    return False
        except FileNotFoundError:
            return True
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при удалении пункта")
            return False
    except Exception:
        messagebox.showerror("Ошибка", "Ошибка при удалении пункта")
        return False