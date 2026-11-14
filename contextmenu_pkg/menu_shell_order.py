"""Порядок пунктов в DesktopBackground\\Shell и CommandStore."""
import winreg

import logger
from core import reg_utils
from contextmenu_pkg.menu_checks import check_menu_item_exists

MOD = "contextmenu"


def update_desktop_shell_order():
    try:
        shell_key = r'DesktopBackground\Shell'
        active_items = []

        if check_menu_item_exists("task_manager"):
            active_items.append("TaskManager")
        if check_menu_item_exists("restart_explorer"):
            active_items.append("RestartExplorer")
        if check_menu_item_exists("clipboard_clear"):
            active_items.append("Clip")
        if check_menu_item_exists("regedit"):
            active_items.append("Regedit")

        if check_menu_item_exists("power_menu"):
            power_keys = [
                r'DesktopBackground\Shell\Shutdown',
                r'DesktopBackground\Shell\Reboot',
                r'DesktopBackground\Shell\Hibernation'
            ]
            for key_path in power_keys:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                        if "Shutdown" in key_path:
                            active_items.append("Shutdown")
                        elif "Reboot" in key_path:
                            active_items.append("Reboot")
                        elif "Hibernation" in key_path:
                            active_items.append("Hibernation")
                except FileNotFoundError:
                    pass

        active_items.append("Display")
        active_items.append("Personalize")

        order_string = ",".join(active_items)

        logger.log_reg(MOD, "set", shell_key, "", order_string)
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, shell_key, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, order_string)
    except Exception:
        try:
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r'DesktopBackground\Shell')
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'DesktopBackground\Shell',
                                0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Display,Personalize")
        except Exception:
            pass


def create_command_store_entries():
    try:
        command_store_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell'
        entries = {
            'cmd': ('Командная строка', 'cmd.exe', 'cmd.exe'),
            'perfmon': ('Монитор ресурсов', 'imageres.dll,144', 'perfmon.exe /res'),
            'relmon': ('Монитор надежности системы', 'perfmon.exe', 'perfmon.exe /rel'),
            'trouble': ('Устранение неполадок', 'imageres.dll,124',
                        'control /name Microsoft.Troubleshooting'),
            'services': ('Службы', 'filemgmt.dll', 'mmc.exe services.msc'),
            'event': ('Просмотр событий', 'eventvwr.exe', 'eventvwr.exe'),
            'taskschd': ('Планировщик заданий', 'miguiresource.dll,1',
                         'Control schedtasks'),
            'useracc': ('Учетные записи', 'imageres.dll,74',
                        'Control userpasswords'),
            'useracc2': ('Учетные записи (классич.)', 'shell32.dll,111',
                         'Control userpasswords2'),
            'network': ('Сетевые подключения', 'ncpa.cpl', 'control ncpa.cpl'),
            'wf': ('Брандмауэр Windows', 'wscui.cpl,3', 'wf.msc'),
            'run': ('Выполнить', 'shell32.dll,24',
                    'explorer shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'),
            'control': ('Панель управления', 'control.exe', 'control.exe'),
            'sysdm': ('Свойства системы', 'sysdm.cpl', 'control sysdm.cpl'),
            'msconfig': ('Конфигурация системы', 'msconfig.exe', 'msconfig.exe'),
            'regedit': ('Редактор реестра', 'regedit.exe', 'regedit.exe'),
            'taskmgr': ('Диспетчер задач', 'taskmgr.exe', 'taskmgr.exe'),
            'devmgr': ('Диспетчер устройств', 'devmgr.dll,4',
                       'mmc.exe devmgmt.msc'),
            'cleanmgr': ('Очистка диска', 'cleanmgr.exe', 'cleanmgr.exe'),
            'defrag': ('Дефрагментация диска', 'dfrgui.exe', 'dfrgui.exe'),
            'folderoptions': ('Свойства папки', 'explorer.exe', 'control folders'),
            'networksh': ('Центр управления сетями', 'networkexplorer.dll,4',
                          'control /name Microsoft.NetworkAndSharingCenter'),
            'appwiz': ('Программы и компоненты', 'appwiz.cpl',
                       'control appwiz.cpl'),
            'power': ('Электропитание', 'powercfg.cpl', 'control powercfg.cpl'),
            'rstrui': ('Восстановление системы', 'rstrui.exe', 'rstrui.exe'),
            'search': ('Поиск', 'shell32.dll,22',
                       'explorer shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}'),
        }

        ok_count = 0
        total = len(entries)
        for cmd_name, (display_name, icon, command) in entries.items():
            try:
                cmd_key_path = f"{command_store_path}\\{cmd_name}"
                reg_utils.reg_add(cmd_key_path, '', display_name,
                                  'REG_SZ', module=MOD)
                reg_utils.reg_add(cmd_key_path, 'Icon', icon,
                                  'REG_SZ', module=MOD)
                cmd_command_path = f"{cmd_key_path}\\command"
                reg_utils.reg_add(cmd_command_path, '', command,
                                  'REG_SZ', module=MOD)
                ok_count += 1
            except Exception:
                continue
        return ok_count == total
    except Exception:
        return False