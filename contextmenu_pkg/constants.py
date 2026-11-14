"""Константы для контекстного меню."""

TIP_DESKTOP = "Контекстное меню Рабочего стола"
TIP_COMPUTER = "Контекстное меню Компьютер"

DEFENDER_GUID = "{09A47860-11B0-4DA5-AFA5-26D86198A780}"

MODERN_SHARING_PATHS = (
    r'SOFTWARE\Classes\*\shellex\ContextMenuHandlers\ModernSharing',
    r'SOFTWARE\Classes\AllFilesystemObjects\shellex\ContextMenuHandlers\ModernSharing',
    r'SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\ModernSharing',
    r'SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\ModernSharing',
    r'SOFTWARE\Classes\Directory\Background\shellex\ContextMenuHandlers\ModernSharing',
)

PAINT3D_EXTS = ('bmp', 'fbx', 'gif', 'glb', 'jfif', 'jpe', 'jpeg', 'jpg',
                'obj', 'ply', 'png', 'stl', 'tif', 'tiff')

WALLPAPER_EXTS = ('bmp', 'dib', 'gif', 'jfif', 'jpe', 'jpeg', 'jpg',
                  'png', 'tif', 'tiff', 'wdp', 'heic', 'heif', 'avif')

POWERSHELL_CLASSES = ('Directory', 'Directory\\Background', 'Drive')
PINTOHOME_CLASSES = ('Drive', 'Folder', 'Network')

PHOTOS_APPX_CLSID = r'AppX43hnxtbyyps62jhe9sqpdzxn1790zetc'
PHOTOS_SHELL_EDIT = rf'{PHOTOS_APPX_CLSID}\Shell\ShellEdit'
PHOTOS_SHELL_CREATE_VIDEO = rf'{PHOTOS_APPX_CLSID}\Shell\ShellCreateVideo'

DEFAULT_PRINT_BASES = (
    r"SystemFileAssociations\image",
    r"SystemFileAssociations\text",
    r"SystemFileAssociations\.pdf",
    r"SystemFileAssociations\.bmp",
    r"SystemFileAssociations\.jpg",
    r"SystemFileAssociations\.jpeg",
    r"SystemFileAssociations\.png",
    r"SystemFileAssociations\.gif",
    r"SystemFileAssociations\.tif",
    r"SystemFileAssociations\.tiff",
    r"SystemFileAssociations\.wdp",
    r"SystemFileAssociations\.txt",
    r"SystemFileAssociations\.rtf",
    r"SystemFileAssociations\.docx",
    r"SystemFileAssociations\.xlsx",
    r"SystemFileAssociations\.pptx",
    r"Word.Document.12", r"Word.Document.8", r"Word.RTF.8",
    r"Excel.Sheet.12", r"Excel.Sheet.8",
    r"PowerPoint.Show.12", r"PowerPoint.Show.8",
    r"AcroExch.Document.7", r"AcroExch.Document", r"AcroExch.Document.DC",
    r"FoxitReader.Document",
    r"txtfile", r"rtffile", r"docxfile", r"docfile",
    r"batfile", r"cmdfile", r"fonfile", r"htmlfile",
    r"inffile", r"inifile", r"JSEFile", r"otffile", r"pfmfile",
    r"regfile", r"ttcfile", r"ttffile", r"VBEFile", r"VBSFile", r"WSFFile",
)

CONTEXT_MENU_KEYS = {
    'pintohome': 'pintohome',
    'powershell': 'powershell',
    'share_access': 'share_access',
    'modern_sharing': 'modern_sharing',
    'defender': 'defender',
    'previous_version': 'previous_version',
    'pin_to_start': 'pin_to_start',
    'pin_to_taskbar': 'pin_to_taskbar',
    'copy_as_path': 'copy_as_path',
    'compatibility': 'compatibility',
    'paint3d': 'paint3d',
    'photos_edit': 'photos_edit',
    'create_video': 'create_video',
    'wallpaper': 'wallpaper',
    'print': 'print',
    'send_to_device': 'send_to_device',
    'take_ownership': 'take_ownership',
    'copy_move': 'copy_move',
    'delete_folder_content': 'delete_folder_content',
    'task_manager': 'task_manager',
    'clipboard_clear': 'clipboard_clear',
    'restart_explorer': 'restart_explorer',
    'regedit': 'regedit',
    'power_menu': 'power_menu',
    'admin_menu': 'admin_menu',
    'system_menu': 'system_menu',
}

ITEM_NAME_BY_KEY = {
    'pintohome': 'Закрепить на панели быстрого доступа',
    'powershell': 'Открыть окно PowerShell здесь',
    'share_access': 'Предоставить доступ к',
    'modern_sharing': 'Отправить (metro)',
    'defender': 'Проверка с использованием Windows Defender',
    'previous_version': 'Восстановить прежнюю версию',
    'pin_to_start': 'Закрепить на начальном экране',
    'pin_to_taskbar': 'Закрепить на Панели задач',
    'copy_as_path': 'Копировать как путь',
    'compatibility': 'Исправление проблем с совместимостью',
    'paint3d': 'Изменить с помощью Paint 3D',
    'photos_edit': 'Изменить с помощью Фотографии',
    'create_video': 'Создать новое видео',
    'wallpaper': 'Сделать фоном рабочего стола',
    'print': 'Печать',
    'send_to_device': 'Передать на устройство',
    'take_ownership': 'Стать владельцем',
    'copy_move': 'Копировать/Переместить в папку',
    'delete_folder_content': 'Удалить содержимое папки',
    'task_manager': 'Диспетчер задач',
    'clipboard_clear': 'Очистить буфер',
    'restart_explorer': 'Перезапустить explorer',
    'regedit': 'Редактор реестра',
    'power_menu': 'Выключение/Перезагрузка/Гибернация',
    'admin_menu': 'Администрирование',
    'system_menu': 'Система',
}