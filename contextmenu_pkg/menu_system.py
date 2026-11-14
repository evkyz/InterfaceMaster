"""Пункт «Система»."""
import subprocess
from tkinter import messagebox

import logger
from core import reg_utils
from contextmenu_pkg.menu_shell_order import create_command_store_entries

MOD = "contextmenu"


def add_system_menu():
    if not create_command_store_entries():
        if not run_as_admin_add_system_menu():
            messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Система'")
        return

    try:
        clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'
        reg_utils.reg_add(clsid_path_hkcu, 'MUIVerb', 'Система',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(
            clsid_path_hkcu, 'SubCommands',
            "control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;"
            "defrag;folderoptions;networksh;appwiz;power;rstrui;search",
            'REG_SZ', module=MOD)
        reg_utils.reg_add(clsid_path_hkcu, 'Icon', 'imageres.dll,104',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(clsid_path_hkcu, 'Position', 'Top',
                          'REG_SZ', module=MOD)
    except Exception:
        if not run_as_admin_add_system_menu():
            messagebox.showerror("Ошибка", "Не удалось добавить пункт 'Система'")


def run_as_admin_add_system_menu():
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["powershell", "-Command", _SYSTEM_PS],
            capture_output=True, text=True, encoding='cp866',
            timeout=30, startupinfo=startupinfo, shell=True
        )
        return result.returncode == 0
    except Exception:
        return False


def remove_system_menu():
    try:
        import winreg
        clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu)
        except Exception:
            pass
        clsid_path_hklm = r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\System'
        try:
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, clsid_path_hklm)
        except Exception:
            pass
        logger.log_reg(MOD, "delete", clsid_path_hkcu)
        return True
    except Exception:
        success = run_as_admin_remove_system_menu()
        if not success:
            messagebox.showerror("Ошибка", "Не удалось удалить пункт 'Система'")
        return False


def run_as_admin_remove_system_menu():
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["powershell", "-Command", _SYSTEM_REMOVE_PS],
            capture_output=True, text=True, encoding='cp866',
            timeout=30, startupinfo=startupinfo, shell=True
        )
        return result.returncode == 0
    except Exception:
        return False


_SYSTEM_PS = '''
$commandStorePath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell"
$systemCommands = @{
    'control' = @('Панель управления', 'control.exe', 'control.exe')
    'sysdm' = @('Свойства системы', 'sysdm.cpl', 'control sysdm.cpl')
    'msconfig' = @('Конфигурация системы', 'msconfig.exe', 'msconfig.exe')
    'regedit' = @('Редактор реестра', 'regedit.exe', 'regedit.exe')
    'taskmgr' = @('Диспетчер задач', 'taskmgr.exe', 'taskmgr.exe')
    'devmgr' = @('Диспетчер устройств', 'devmgr.dll,4', 'mmc.exe devmgmt.msc')
    'cleanmgr' = @('Очистка диска', 'cleanmgr.exe', 'cleanmgr.exe')
    'defrag' = @('Дефрагментация диска', 'dfrgui.exe', 'dfrgui.exe')
    'folderoptions' = @('Свойства папки', 'explorer.exe', 'control folders')
    'networksh' = @('Центр управления сетями', 'networkexplorer.dll,4', 'control /name Microsoft.NetworkAndSharingCenter')
    'appwiz' = @('Программы и компоненты', 'appwiz.cpl', 'control appwiz.cpl')
    'power' = @('Электропитание', 'powercfg.cpl', 'control powercfg.cpl')
    'rstrui' = @('Восстановление системы', 'rstrui.exe', 'rstrui.exe')
    'search' = @('Поиск', 'shell32.dll,22', 'explorer shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}')
}
foreach ($cmdName in $systemCommands.Keys) {
    $displayName, $icon, $command = $systemCommands[$cmdName]
    $cmdKeyPath = "$commandStorePath\\$cmdName"
    try {
        New-Item -Path $cmdKeyPath -Force | Out-Null
        Set-ItemProperty -Path $cmdKeyPath -Name "(default)" -Value $displayName -Type String
        Set-ItemProperty -Path $cmdKeyPath -Name "Icon" -Value $icon -Type String
        $cmdCommandPath = "$cmdKeyPath\\command"
        New-Item -Path $cmdCommandPath -Force | Out-Null
        Set-ItemProperty -Path $cmdCommandPath -Name "(default)" -Value $command -Type String
    } catch { }
}
$clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"
try {
    New-Item -Path $clsidPathHKCU -Force -ErrorAction Stop | Out-Null
    Set-ItemProperty -Path $clsidPathHKCU -Name "MUIVerb" -Value "Система" -Type String
    Set-ItemProperty -Path $clsidPathHKCU -Name "SubCommands" -Value "control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;defrag;folderoptions;networksh;appwiz;power;rstrui;search" -Type String
    Set-ItemProperty -Path $clsidPathHKCU -Name "Icon" -Value "imageres.dll,104" -Type String
    Set-ItemProperty -Path $clsidPathHKCU -Name "Position" -Value "Top" -Type String
    return $true
} catch {
    $clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"
    try {
        if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
            return $false
        }
        New-Item -Path $clsidPathHKLM -Force -ErrorAction Stop | Out-Null
        Set-ItemProperty -Path $clsidPathHKLM -Name "MUIVerb" -Value "Система" -Type String
        Set-ItemProperty -Path $clsidPathHKLM -Name "SubCommands" -Value "control;sysdm;msconfig;regedit;taskmgr;devmgr;cleanmgr;defrag;folderoptions;networksh;appwiz;power;rstrui;search" -Type String
        Set-ItemProperty -Path $clsidPathHKLM -Name "Icon" -Value "imageres.dll,104" -Type String
        Set-ItemProperty -Path $clsidPathHKLM -Name "Position" -Value "Top" -Type String
        return $true
    } catch { return $false }
}
'''

_SYSTEM_REMOVE_PS = '''
$clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"
$clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\System"
$removed = $false
try { if (Test-Path $clsidPathHKCU) { Remove-Item -Path $clsidPathHKCU -Recurse -Force -ErrorAction Stop; $removed = $true } } catch { }
try { if (Test-Path $clsidPathHKLM) { Remove-Item -Path $clsidPathHKLM -Recurse -Force -ErrorAction Stop; $removed = $true } } catch { }
return $removed
'''