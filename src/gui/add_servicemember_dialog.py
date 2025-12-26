"""
Діалог для додавання нового військовослужбовця
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QMessageBox,
    QGroupBox, QCompleter
)
from PySide6.QtCore import Qt


class AddServicememberDialog(QDialog):
    """
    Діалог для додавання нового військовослужбовця в БД
    """

    def __init__(self, db_manager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.added_id = None  # ID доданого військовослужбовця
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Додати військовослужбовця")
        self.setMinimumWidth(450)

        layout = QVBoxLayout()

        # Стилі для полів введення
        combo_style = "QComboBox { color: black; background-color: white; }"
        input_style = "QLineEdit { color: black; background-color: white; }"

        # Група полів
        group = QGroupBox("Дані військовослужбовця")
        group_layout = QVBoxLayout()

        # ПІБ
        self.name_input = self.create_field(group_layout, "ПІБ (Прізвище Ім'я По-батькові):", required=True, style=input_style)

        # Звання (dropdown з можливістю введення)
        rank_layout = QHBoxLayout()
        rank_layout.addWidget(QLabel("Звання:"))
        self.rank_combo = QComboBox()
        self.rank_combo.setEditable(True)
        self.rank_combo.setInsertPolicy(QComboBox.NoInsert)
        self.rank_combo.setStyleSheet(combo_style)

        # Заповнюємо список звань з БД
        try:
            ranks = self.db_manager.get_unique_ranks()
            if ranks:
                self.rank_combo.addItems(ranks)
            else:
                # Стандартні звання якщо БД порожня
                default_ranks = [
                    "солдат", "старший солдат", "молодший сержант",
                    "сержант", "старший сержант", "головний сержант",
                    "штаб-сержант", "майстер-сержант", "старший майстер-сержант",
                    "головний майстер-сержант", "молодший лейтенант", "лейтенант",
                    "старший лейтенант", "капітан", "майор", "підполковник", "полковник"
                ]
                self.rank_combo.addItems(default_ranks)
        except:
            pass
        rank_layout.addWidget(self.rank_combo)
        group_layout.addLayout(rank_layout)

        # Посада
        self.position_input = self.create_field(group_layout, "Посада:", style=input_style)

        # Підрозділ (dropdown з можливістю введення)
        unit_layout = QHBoxLayout()
        unit_layout.addWidget(QLabel("Підрозділ:"))
        self.unit_combo = QComboBox()
        self.unit_combo.setEditable(True)
        self.unit_combo.setInsertPolicy(QComboBox.NoInsert)
        self.unit_combo.setStyleSheet(combo_style)

        # Заповнюємо список підрозділів з БД
        try:
            units = self.db_manager.get_unique_units()
            if units:
                self.unit_combo.addItems(units)
        except:
            pass

        # Додаємо автодоповнення
        completer = QCompleter(self.unit_combo.model())
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.unit_combo.setCompleter(completer)

        unit_layout.addWidget(self.unit_combo)
        group_layout.addLayout(unit_layout)

        # РНОКПП
        self.rnokpp_input = self.create_field(group_layout, "РНОКПП:", style=input_style)
        self.rnokpp_input.setPlaceholderText("10 цифр")

        # Дата народження
        birth_layout = QHBoxLayout()
        birth_layout.addWidget(QLabel("Дата народження:"))
        self.birth_date_input = QLineEdit()
        self.birth_date_input.setPlaceholderText("01.01.1990")
        self.birth_date_input.setStyleSheet(input_style)
        birth_layout.addWidget(self.birth_date_input)
        group_layout.addLayout(birth_layout)

        group.setLayout(group_layout)
        layout.addWidget(group)

        # Кнопки
        button_layout = QHBoxLayout()

        add_btn = QPushButton("Додати")
        add_btn.setStyleSheet("background-color: #4caf50; color: white; font-weight: bold; padding: 8px 16px;")
        add_btn.clicked.connect(self.add_servicemember)
        button_layout.addWidget(add_btn)

        cancel_btn = QPushButton("Скасувати")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def create_field(self, layout, label_text, required=False, style=None):
        """Створює поле введення з міткою"""
        field_layout = QHBoxLayout()
        label = QLabel(label_text)
        if required:
            label.setStyleSheet("font-weight: bold;")
        field_layout.addWidget(label)
        input_field = QLineEdit()
        if style:
            input_field.setStyleSheet(style)
        field_layout.addWidget(input_field)
        layout.addLayout(field_layout)
        return input_field

    def add_servicemember(self):
        """Додає нового військовослужбовця в БД"""
        # Перевірка обов'язкових полів
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Помилка", "ПІБ є обов'язковим полем!")
            self.name_input.setFocus()
            return

        # Перевіряємо чи існує
        existing = self.db_manager.get_servicemember_by_name(name)
        if existing:
            reply = QMessageBox.question(
                self,
                "Увага",
                f"Військовослужбовець '{name}' вже існує в базі.\n"
                "Оновити його дані?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        # Збираємо дані
        data = {
            "name": name,
            "rank": self.rank_combo.currentText().strip(),
            "position": self.position_input.text().strip(),
            "rnokpp": self.rnokpp_input.text().strip(),
            "unit": self.unit_combo.currentText().strip(),
            "birth_date": self.birth_date_input.text().strip()
        }

        try:
            # Додаємо в БД
            self.added_id = self.db_manager.add_servicemember(data)

            QMessageBox.information(
                self,
                "Успіх",
                f"Військовослужбовця '{name}' успішно додано!"
            )

            self.accept()

        except Exception as e:
            QMessageBox.critical(
                self,
                "Помилка",
                f"Не вдалось додати військовослужбовця:\n{str(e)}"
            )
