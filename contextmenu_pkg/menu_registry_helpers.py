"""Хелперы для работы с реестром: rename/copy/delete/take_ownership/set_permissions."""
import os
import subprocess
import winreg

import logger

MOD = "contextmenu"


def set_registry_key_permissions_simple(key_path):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["reg", "add", f"HKLM\\{key_path}", "/f"],
            capture_output=True, text=True, encoding='cp866',
            timeout=10, startupinfo=startupinfo, shell=True
        )
        if result.returncode != 0:
            ps_command = f'''
            try {{
                $key = "HKLM:\\{key_path}"
                New-Item -Path $key -Force -ErrorAction Stop | Out-Null
                return $true
            }} catch {{ return $false }}
            '''
            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True, text=True, encoding='cp866',
                timeout=10, startupinfo=startupinfo, shell=True
            )
            return result.returncode == 0
        return True
    except Exception:
        return False


def take_ownership_registry_key_simple(key_path):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["reg", "add", f"HKLM\\{key_path}", "/f", "/reg:64"],
            capture_output=True, text=True, encoding='cp866',
            timeout=10, startupinfo=startupinfo, shell=True
        )
        if result.returncode != 0:
            username = os.environ.get("USERNAME", "")
            ps_command = f'''
            try {{
                $key = "HKLM:\\{key_path}"
                $acl = Get-Acl $key -ErrorAction Stop
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("BUILTIN\\Administrators","FullControl","Allow")
                $acl.SetAccessRule($rule)
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("NT AUTHORITY\\SYSTEM","FullControl","Allow")
                $acl.SetAccessRule($rule)
                $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("{username}","FullControl","Allow")
                $acl.SetAccessRule($rule)
                Set-Acl $key $acl -ErrorAction Stop
                return $true
            }} catch {{
                try {{
                    New-Item -Path "HKLM:\\{key_path}" -Force -ErrorAction Stop | Out-Null
                    return $true
                }} catch {{ return $false }}
            }}
            '''
            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True, text=True, encoding='cp866',
                timeout=10, startupinfo=startupinfo, shell=True
            )
            return result.returncode == 0
        return True
    except Exception:
        return False


def set_registry_key_permissions(hive, key_path):
    try:
        if hive == winreg.HKEY_CLASSES_ROOT:
            root = "HKCR"
        elif hive == winreg.HKEY_CURRENT_USER:
            root = "HKCU"
        else:
            root = "HKLM"
        full_path = f"{root}\\{key_path}"
        username = os.environ.get("USERNAME", "")
        ps_command = f'''
        $key = "Registry::{full_path}"
        try {{
            $acl = Get-Acl $key
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("BUILTIN\\Administrators","FullControl","Allow")
            $acl.SetAccessRule($rule)
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("NT AUTHORITY\\SYSTEM","FullControl","Allow")
            $acl.SetAccessRule($rule)
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("{username}","FullControl","Allow")
            $acl.SetAccessRule($rule)
            Set-Acl $key $acl
        }} catch {{ }}
        '''
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True, text=True, encoding='cp866',
            timeout=30, startupinfo=startupinfo
        )
        return result.returncode == 0
    except Exception:
        return False


def take_ownership_registry_key(hive, key_path):
    try:
        if hive == winreg.HKEY_CLASSES_ROOT:
            root = "HKCR"
        elif hive == winreg.HKEY_CURRENT_USER:
            root = "HKCU"
        else:
            root = "HKLM"
        full_path = f"{root}\\{key_path}"
        ps_command = f'''
        $key = "Registry::{full_path}"
        try {{
            $acl = Get-Acl $key
            $owner = [System.Security.Principal.NTAccount]"BUILTIN\\Administrators"
            $acl.SetOwner($owner)
            Set-Acl $key $acl
        }} catch {{ }}
        '''
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True, text=True, encoding='cp866',
            timeout=30, startupinfo=startupinfo
        )
        return result.returncode == 0
    except Exception:
        return False


def delete_registry_key_recursive(hive, key_path):
    try:
        with winreg.OpenKey(hive, key_path) as key:
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    delete_registry_key_recursive(hive, key_path + "\\" + subkey_name)
                except OSError:
                    break
                i += 1
    except FileNotFoundError:
        return
    except Exception:
        pass
    try:
        winreg.DeleteKey(hive, key_path)
    except Exception:
        pass


def copy_registry_key(hive, source_path, dest_path):
    winreg.CreateKey(hive, dest_path)
    with winreg.OpenKey(hive, source_path) as source_key:
        with winreg.OpenKey(hive, dest_path, 0, winreg.KEY_WRITE) as dest_key:
            i = 0
            while True:
                try:
                    name, value, type_val = winreg.EnumValue(source_key, i)
                    winreg.SetValueEx(dest_key, name, 0, type_val, value)
                    i += 1
                except OSError:
                    break
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(source_key, i)
                    copy_registry_key(hive, source_path + "\\" + subkey_name,
                                      dest_path + "\\" + subkey_name)
                    i += 1
                except OSError:
                    break
    return True


def delete_registry_key(hive, key_path):
    try:
        try:
            with winreg.OpenKey(hive, key_path) as key:
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        delete_registry_key(hive, key_path + "\\" + subkey_name)
                        i += 1
                    except OSError:
                        break
        except FileNotFoundError:
            return True
        winreg.DeleteKey(hive, key_path)
        return True
    except PermissionError:
        raise
    except FileNotFoundError:
        return True
    except Exception:
        raise


def rename_registry_key(hive, old_key_path, new_key_path):
    for _ in range(3):
        try:
            copy_registry_key(hive, old_key_path, new_key_path)
            delete_registry_key(hive, old_key_path)
            return True
        except PermissionError:
            if take_ownership_registry_key_simple(old_key_path):
                if set_registry_key_permissions_simple(old_key_path):
                    try:
                        copy_registry_key(hive, old_key_path, new_key_path)
                        delete_registry_key(hive, old_key_path)
                        return True
                    except PermissionError:
                        continue
        except FileNotFoundError:
            try:
                with winreg.OpenKey(hive, new_key_path):
                    return True
            except FileNotFoundError:
                return False
        except Exception:
            continue
    return False