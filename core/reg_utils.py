"""Работа с реестром: reg.exe + winreg + логирование."""
import subprocess
import winreg

import logger

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

_HIVE_MAP = {
    'HKCU': winreg.HKEY_CURRENT_USER,
    'HKLM': winreg.HKEY_LOCAL_MACHINE,
    'HKCR': winreg.HKEY_CLASSES_ROOT,
}


def reg(*args, module=None, log=True):
    """subprocess reg с опциональным логированием add/delete."""
    if log and module and args and args[0] in ('add', 'delete'):
        action = args[0]
        key = args[1] if len(args) > 1 else ''
        value_name = None
        value_data = None
        i = 2
        while i < len(args):
            a = args[i]
            if a == '/ve':
                value_name = ''
                i += 1
            elif a == '/v':
                value_name = args[i + 1]
                i += 2
            elif a == '/d':
                value_data = args[i + 1]
                i += 2
            else:
                i += 1
        logger.log_reg(module, action, key, value_name, value_data)
    return subprocess.run(
        ['reg', *args], capture_output=True, text=True,
        encoding='cp866', creationflags=CREATE_NO_WINDOW)


def reg_add(hive_path, value_name=None, value_data=None,
            value_type='REG_SZ', module=None, log=True):
    """reg add с удобным API.

    hive_path — полный путь, напр. 'HKCU\\Software\\...'.
    value_name — None = default (/ve), строка = /v имя.
    """
    args = ['add', hive_path]
    if value_name is None:
        args.append('/ve')
    else:
        args.extend(['/v', value_name])
    if value_data is not None:
        args.extend(['/d', str(value_data)])
    if value_type:
        args.extend(['/t', value_type])
    args.append('/f')
    return reg(*args, module=module, log=log)


def reg_delete(hive_path, value_name=None, module=None, log=True):
    args = ['delete', hive_path]
    if value_name is not None:
        args.extend(['/v', value_name])
    args.append('/f')
    return reg(*args, module=module, log=log)


def reg_query(hive_path, value_name=None):
    args = ['query', hive_path]
    if value_name is not None:
        args.extend(['/v', value_name])
    return reg(*args, log=False)


def get_value(hive, key_path, value_name=''):
    """winreg-чтение значения. Возвращает None, если ключа/значения нет."""
    try:
        with winreg.OpenKey(_HIVE_MAP[hive], key_path) as k:
            value, vtype = winreg.QueryValueEx(k, value_name)
            if vtype == winreg.REG_DWORD:
                return hex(value)
            return value
    except (FileNotFoundError, OSError, KeyError):
        return None


def key_exists(hive, key_path):
    try:
        with winreg.OpenKey(_HIVE_MAP[hive], key_path):
            return True
    except (FileNotFoundError, OSError, KeyError):
        return False


def set_value(hive, key_path, value_name, value_data, value_type=winreg.REG_SZ,
              module=None):
    """winreg-запись с логированием."""
    try:
        try:
            key = winreg.OpenKey(_HIVE_MAP[hive], key_path, 0, winreg.KEY_WRITE)
        except FileNotFoundError:
            key = winreg.CreateKey(_HIVE_MAP[hive], key_path)
        winreg.SetValueEx(key, value_name, 0, value_type, value_data)
        winreg.CloseKey(key)
        if module:
            logger.log_reg(module, "set", f"{hive}\\{key_path}", value_name, value_data)
        return True
    except (PermissionError, OSError):
        return False


def delete_value(hive, key_path, value_name, module=None):
    try:
        with winreg.OpenKey(_HIVE_MAP[hive], key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, value_name)
        if module:
            logger.log_reg(module, "delete", f"{hive}\\{key_path}", value_name)
        return True
    except FileNotFoundError:
        return True
    except OSError:
        return False

def reg_create_key(hive_path, module=None, log=True):
    """Создать ключ (reg add key /f), без значений."""
    if log and module:
        logger.log_reg(module, "add", hive_path)
    return subprocess.run(
        ['reg', 'add', hive_path, '/f'], capture_output=True, text=True,
        encoding='cp866', creationflags=CREATE_NO_WINDOW)