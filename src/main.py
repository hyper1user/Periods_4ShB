"""
Точка входу в додаток
"""
import sys
import os

# Для PyInstaller: визначаємо базовий шлях
if getattr(sys, 'frozen', False):
    # Запущено як exe
    base_path = sys._MEIPASS
else:
    # Запущено як скрипт
    base_path = os.path.dirname(os.path.abspath(__file__))

# Додаємо шлях до модулів
sys.path.insert(0, base_path)

from PySide6.QtWidgets import QApplication
from gui.main_window import MainWindow
from gui.styles import get_military_style


def main():
    """
    Головна функція додатку
    """
    # Створення додатку
    app = QApplication(sys.argv)

    # Застосування стилю
    app.setStyleSheet(get_military_style())

    # Створення та показ головного вікна
    window = MainWindow()
    window.show()

    # Запуск циклу подій
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
