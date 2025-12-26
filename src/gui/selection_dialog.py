"""
Діалогове вікно для вибору військовослужбовців
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QComboBox, QCheckBox, QPushButton, QButtonGroup, QRadioButton,
    QGroupBox
)
from PySide6.QtCore import Qt
from typing import List, Optional


class SelectionDialog(QDialog):
    """
    Діалог для вибору цільової аудиторії генерації рапортів
    """

    def __init__(self, names: List[str], units: List[str] = None, parent=None):
        """
        Ініціалізація діалогу

        Args:
            names: Список ПІБ військовослужбовців
            units: Список підрозділів (опціонально)
            parent: Батьківський віджет
        """
        super().__init__(parent)
        self.names = names
        self.units = units or []
        self.selected_name = None
        self.selected_unit = None
        self.selection_mode = "single"  # single, all, unit

        self.init_ui()

    def init_ui(self):
        """
        Ініціалізація інтерфейсу
        """
        self.setWindowTitle("Вибір військовослужбовців")
        self.setMinimumWidth(500)
        self.setMinimumHeight(300)

        layout = QVBoxLayout()

        # Заголовок
        title = QLabel("Оберіть для кого створити рапорти:")
        title.setObjectName("title")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        # Група вибору режиму
        mode_group = QGroupBox("Режим вибору")
        mode_layout = QVBoxLayout()

        # Радіокнопки для вибору режиму
        self.radio_single = QRadioButton("Один військовослужбовець")
        self.radio_all = QRadioButton("Створити на всіх")
        self.radio_unit = QRadioButton("Створити на підрозділ")

        self.radio_single.setChecked(True)
        self.radio_single.toggled.connect(self.on_mode_changed)
        self.radio_all.toggled.connect(self.on_mode_changed)
        self.radio_unit.toggled.connect(self.on_mode_changed)

        mode_layout.addWidget(self.radio_single)
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_unit)

        mode_group.setLayout(mode_layout)
        layout.addWidget(mode_group)

        # Випадаючий список для вибору ПІБ з автодоповненням
        self.name_label = QLabel("Оберіть військовослужбовця:")
        layout.addWidget(self.name_label)

        self.name_combo = QComboBox()
        self.name_combo.setEditable(True)  # Дозволяємо редагування
        self.name_combo.setInsertPolicy(QComboBox.NoInsert)  # Не додаємо нові елементи
        self.name_combo.addItems(self.names)
        # Виправлення стилю для видимості тексту
        self.name_combo.setStyleSheet("QComboBox { color: black; background-color: white; }")

        # Додаємо автодоповнення
        from PySide6.QtWidgets import QCompleter
        from PySide6.QtCore import Qt
        completer = QCompleter(self.names)
        completer.setCaseSensitivity(Qt.CaseInsensitive)  # Ігноруємо регістр
        completer.setFilterMode(Qt.MatchContains)  # Шукаємо у будь-якій частині рядка
        # Виправлення стилю для автодоповнення
        completer.popup().setStyleSheet("QListView { color: black; background-color: white; }")
        self.name_combo.setCompleter(completer)

        layout.addWidget(self.name_combo)

        # Випадаючий список для вибору підрозділу з автодоповненням
        self.unit_label = QLabel("Оберіть підрозділ:")
        self.unit_label.setVisible(False)
        layout.addWidget(self.unit_label)

        self.unit_combo = QComboBox()
        self.unit_combo.setEditable(True)
        self.unit_combo.setInsertPolicy(QComboBox.NoInsert)
        # Виправлення стилю для видимості тексту
        self.unit_combo.setStyleSheet("QComboBox { color: black; background-color: white; }")
        if self.units:
            self.unit_combo.addItems(self.units)
            unit_completer = QCompleter(self.units)
            unit_completer.setCaseSensitivity(Qt.CaseInsensitive)
            unit_completer.setFilterMode(Qt.MatchContains)
            # Виправлення стилю для автодоповнення
            unit_completer.popup().setStyleSheet("QListView { color: black; background-color: white; }")
            self.unit_combo.setCompleter(unit_completer)
        self.unit_combo.setVisible(False)
        layout.addWidget(self.unit_combo)

        # Інформаційна мітка
        self.info_label = QLabel("")
        self.info_label.setStyleSheet("color: gray; font-size: 12px;")
        layout.addWidget(self.info_label)
        self.update_info_label()

        # Кнопки
        button_layout = QHBoxLayout()

        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)

        self.cancel_button = QPushButton("Скасувати")
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)

        layout.addStretch()
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def on_mode_changed(self):
        """
        Обробник зміни режиму вибору
        """
        if self.radio_single.isChecked():
            self.selection_mode = "single"
            self.name_label.setVisible(True)
            self.name_combo.setVisible(True)
            self.unit_label.setVisible(False)
            self.unit_combo.setVisible(False)
        elif self.radio_all.isChecked():
            self.selection_mode = "all"
            self.name_label.setVisible(False)
            self.name_combo.setVisible(False)
            self.unit_label.setVisible(False)
            self.unit_combo.setVisible(False)
        elif self.radio_unit.isChecked():
            self.selection_mode = "unit"
            self.name_label.setVisible(False)
            self.name_combo.setVisible(False)
            self.unit_label.setVisible(True)
            self.unit_combo.setVisible(True)

        self.update_info_label()

    def update_info_label(self):
        """
        Оновлення інформаційної мітки
        """
        if self.selection_mode == "single":
            self.info_label.setText(f"Буде створено 1 документ")
        elif self.selection_mode == "all":
            count = len(self.names)
            self.info_label.setText(f"Буде створено {count} документів")
        elif self.selection_mode == "unit":
            self.info_label.setText("Буде створено документи для обраного підрозділу")

    def get_selection(self):
        """
        Отримати результат вибору

        Returns:
            Кортеж (режим, значення)
            - ("single", "ПІБ")
            - ("all", None)
            - ("unit", "Назва підрозділу")
        """
        if self.selection_mode == "single":
            return ("single", self.name_combo.currentText())
        elif self.selection_mode == "all":
            return ("all", None)
        elif self.selection_mode == "unit":
            return ("unit", self.unit_combo.currentText())

        return (None, None)
