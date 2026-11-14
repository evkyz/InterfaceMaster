"""Вспомогательные функции: imaster.dll, admin, debug."""
import os
import sys

from core import win_utils
import logger

MOD = "contextmenu"


def get_imaster_path():
    """Путь к imaster.dll без запроса копирования. None — если не найден
    или лежит внутри _MEIPASS (тогда иконки брать бессмысленно)."""
    try:
        from main import DLLChecker
        path = DLLChecker.get_dll_path_without_prompt()
        if not path or not os.path.isfile(path):
            return None
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            try:
                if os.path.commonpath([os.path.abspath(path),
                                       os.path.abspath(meipass)]) == \
                        os.path.abspath(meipass):
                    return None
            except ValueError:
                pass
        return path
    except Exception:
        pass
    return None


def is_admin():
    return win_utils.is_admin()