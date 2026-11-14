"""Пункт «Администрирование»."""
import subprocess
from tkinter import messagebox

import logger
from core import reg_utils, win_utils
from contextmenu_pkg.menu_shell_order import create_command_store_entries

MOD = "contextmenu"


def add_admin_menu():
    if not create_command_store_entries():
        if not run_as_admin_add_admin_menu():
            messagebox.showerror("Ошибка",
                                 "Не удалось добавить пункт 'Администрирование'")
        return

    try:
        clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'
        reg_utils.reg_add(clsid_path_hkcu, 'MUIVerb', 'Администрирование',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(
            clsid_path_hkcu, 'SubCommands',
            "cmd;perfmon;relmon;trouble;services;event;taskschd;"
            "useracc;useracc2;network;wf;run", 'REG_SZ', module=MOD)
        reg_utils.reg_add(clsid_path_hkcu, 'Icon', 'mmc.exe',
                          'REG_SZ', module=MOD)
        reg_utils.reg_add(clsid_path_hkcu, 'Position', 'Top',
                          'REG_SZ', module=MOD)
    except Exception:
        if not run_as_admin_add_admin_menu():
            messagebox.showerror("Ошибка",
                                 "Не удалось добавить пункт 'Администрирование'")


def run_as_admin_add_admin_menu():
    try:
        ps_command = _ADMIN_PS
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True, text=True, encoding='cp866',
            timeout=30, startupinfo=startupinfo, shell=True
        )
        return result.returncode == 0
    except Exception:
        return False


def remove_admin_menu():
    try:
        clsid_path_hkcu = r'Software\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'
        try:
            import winreg
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, clsid_path_hkcu)
        except Exception:
            pass
        clsid_path_hklm = r'SOFTWARE\Classes\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\shell\Admin'
        try:
            import winreg
            winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, clsid_path_hklm)
        except Exception:
            pass
        logger.log_reg(MOD, "delete", clsid_path_hkcu)
        return True
    except Exception:
        success = run_as_admin_remove_admin_menu()
        if not success:
            messagebox.showerror("Ошибка",
                                 "Не удалось удалить пункт 'Администрирование'")
        return False


def run_as_admin_remove_admin_menu():
    try:
        ps_command = _ADMIN_REMOVE_PS
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True, text=True, encoding='cp866',
            timeout=30, startupinfo=startupinfo, shell=True
        )
        return result.returncode == 0
    except Exception:
        return False


_ADMIN_PS = '''
$commandStorePath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\CommandStore\\shell"
$adminCommands = @{
    'cmd' = @('Командная строка', 'cmd.exe', 'cmd.exe')
    'perfmon' = @('Монитор ресурсов', 'imageres.dll,144', 'perfmon.exe /res')
    'relmon' = @('Монитор надежности системы', 'perfmon.exe', 'perfmon.exe /rel')
    'trouble' = @('Устранение неполадок', 'imageres.dll,124', 'control /name Microsoft.Troubleshooting')
    'services' = @('Службы', 'filemgmt.dll', 'mmc.exe services.msc')
    'event' = @('Просмотр событий', 'eventvwr.exe', 'eventvwr.exe')
    'taskschd' = @('Планировщик заданий', 'miguiresource.dll,1', 'Control schedtasks')
    'useracc' = @('Учетные записи', 'imageres.dll,74', 'Control userpasswords')
    'useracc2' = @('Учетные записи (классич.)', 'shell32.dll,111', 'Control userpasswords2')
    'network' = @('Сетевые подключения', 'ncpa.cpl', 'control ncpa.cpl')
    'wf' = @('Брандмауэр Windows', 'wscui.cpl,3', 'wf.msc')
    'run' = @('Выполнить', 'shell32.dll,24', 'explorer shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}')
}
foreach ($cmdName in $adminCommands.Keys) {
    $displayName, $icon, $command = $adminCommands[$cmdName]
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
$clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"
try {
    New-Item -Path $clsidPathHKCU -Force -ErrorAction Stop | Out-Null
    Set-ItemProperty -Path $clsidPathHKCU -Name "MUIVerb" -Value "Администрирование" -Type String
    Set-ItemProperty -Path $clsidPathHKCU -Name "SubCommands" -Value "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run" -Type String
    Set-ItemProperty -Path $clsidPathHKCU -Name "Icon" -Value "mmc.exe" -Type String
    Set-ItemProperty -Path $clsidPathHKCU -Name "Position" -Value "Top" -Type String
    return $true
} catch {
    $clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"
    try {
        if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
            return $false
        }
        New-Item -Path $clsidPathHKLM -Force -ErrorAction Stop | Out-Null
        Set-ItemProperty -Path $clsidPathHKLM -Name "MUIVerb" -Value "Администрирование" -Type String
        Set-ItemProperty -Path $clsidPathHKLM -Name "SubCommands" -Value "cmd;perfmon;relmon;trouble;services;event;taskschd;useracc;useracc2;network;wf;run" -Type String
        Set-ItemProperty -Path $clsidPathHKLM -Name "Icon" -Value "mmc.exe" -Type String
        Set-ItemProperty -Path $clsidPathHKLM -Name "Position" -Value "Top" -Type String
        return $true
    } catch { return $false }
}
'''

_ADMIN_REMOVE_PS = '''
$clsidPathHKCU = "HKCU:\\Software\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"
$clsidPathHKLM = "HKLM:\\SOFTWARE\\Classes\\CLSID\\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\shell\\Admin"
$removed = $false
try { if (Test-Path $clsidPathHKCU) { Remove-Item -Path $clsidPathHKCU -Recurse -Force -ErrorAction Stop; $removed = $true } } catch { }
try { if (Test-Path $clsidPathHKLM) { Remove-Item -Path $clsidPathHKLM -Recurse -Force -ErrorAction Stop; $removed = $true } } catch { }
return $removed
'''