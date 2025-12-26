"""
Утиліти для роботи зі шляхами в PyInstaller
"""
import sys
import os


def get_base_dir():
    """
    Отримує базову директорію програми для даних користувача
    Використовується для data.db, output та інших файлів користувача

    ВИПРАВЛЕНО: Якщо встановлено в Program Files, використовуємо AppData
    """
    if getattr(sys, 'frozen', False):
        # Запущено як exe (PyInstaller)
        exe_dir = os.path.dirname(sys.executable)

        # Перевірка чи встановлено в Program Files (немає прав на запис)
        program_files = os.environ.get('ProgramFiles', 'C:\\Program Files')
        program_files_x86 = os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)')

        if exe_dir.startswith(program_files) or exe_dir.startswith(program_files_x86):
            # Використовуємо AppData для даних користувача
            app_data = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'Periods_4SHB')
            os.makedirs(app_data, exist_ok=True)
            return app_data
        else:
            # Портативна версія або встановлено в іншу папку - використовуємо папку exe
            return exe_dir
    else:
        # Запущено як скрипт - йдемо з src/utils до кореня проекту
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_resources_dir():
    """
    Отримує директорію ресурсів (templates, config)
    В PyInstaller це _internal папка
    """
    if getattr(sys, 'frozen', False):
        # Запущено як exe - ресурси в _internal (sys._MEIPASS)
        return sys._MEIPASS
    else:
        # Запущено як скрипт - ресурси в корені проекту
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_config_path():
    """Повертає шлях до config/settings.json"""
    return os.path.join(get_resources_dir(), "config", "settings.json")


def get_template_path(template_filename):
    """Повертає шлях до шаблону документа"""
    return os.path.join(get_resources_dir(), "templates", template_filename)


def get_database_path(db_filename="data.db"):
    """Повертає шлях до бази даних"""
    return os.path.join(get_base_dir(), db_filename)


def get_output_dir(output_dirname="output"):
    """Повертає шлях до папки виводу"""
    return os.path.join(get_base_dir(), output_dirname)
