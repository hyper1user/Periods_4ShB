"""
Діалог для редагування та видалення періодів
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QMessageBox,
    QGroupBox, QCompleter, QScrollArea, QWidget,
    QFrame, QInputDialog
)
from PySide6.QtCore import Qt
from datetime import datetime


class EditPeriodsDialog(QDialog):
    """
    Діалог для перегляду, редагування та видалення періодів військовослужбовця
    """

    def __init__(self, db_manager, parent=None, preselected_name=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.preselected_name = preselected_name
        self.current_servicemember_id = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Редагування періодів")
        self.setMinimumWidth(600)
        self.setMinimumHeight(500)

        layout = QVBoxLayout()

        # Вибір військовослужбовця
        select_layout = QHBoxLayout()
        select_layout.addWidget(QLabel("Військовослужбовець:"))

        self.name_combo = QComboBox()
        self.name_combo.setEditable(True)
        self.name_combo.setInsertPolicy(QComboBox.NoInsert)
        self.name_combo.setMinimumWidth(300)
        self.name_combo.setStyleSheet("QComboBox { color: black; background-color: white; }")

        # Заповнюємо список з БД
        try:
            names = self.db_manager.get_unique_names()
            self.name_combo.addItems(names)

            # Автодоповнення
            completer = QCompleter(names)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            completer.setFilterMode(Qt.MatchContains)
            completer.popup().setStyleSheet("QListView { color: black; background-color: white; }")
            self.name_combo.setCompleter(completer)

            # Якщо є preselected_name - вибираємо його
            if self.preselected_name:
                index = self.name_combo.findText(self.preselected_name)
                if index >= 0:
                    self.name_combo.setCurrentIndex(index)
        except:
            pass

        self.name_combo.currentTextChanged.connect(self.load_periods)
        select_layout.addWidget(self.name_combo)

        load_btn = QPushButton("Завантажити")
        load_btn.clicked.connect(self.load_periods)
        select_layout.addWidget(load_btn)

        layout.addLayout(select_layout)

        # Scroll area для періодів
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        self.periods_container = QWidget()
        self.periods_layout = QVBoxLayout()
        self.periods_layout.setAlignment(Qt.AlignTop)
        self.periods_container.setLayout(self.periods_layout)

        scroll.setWidget(self.periods_container)
        layout.addWidget(scroll)

        # Кнопки внизу
        button_layout = QHBoxLayout()

        add_period_btn = QPushButton("+ Додати період")
        add_period_btn.setStyleSheet("background-color: #4caf50; color: white; padding: 8px 16px;")
        add_period_btn.clicked.connect(self.add_new_period)
        button_layout.addWidget(add_period_btn)

        button_layout.addStretch()

        close_btn = QPushButton("Закрити")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

        # Завантажуємо періоди
        if self.preselected_name:
            self.load_periods()

    def load_periods(self):
        """Завантажує періоди для вибраного військовослужбовця"""
        # Очищаємо попередні
        while self.periods_layout.count():
            item = self.periods_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        name = self.name_combo.currentText().strip()
        if not name:
            return

        sm = self.db_manager.get_servicemember_by_name(name)
        if not sm:
            QMessageBox.warning(self, "Помилка", f"Військовослужбовця '{name}' не знайдено!")
            return

        self.current_servicemember_id = sm["id"]

        # Отримуємо періоди
        periods = self.db_manager.get_servicemember_periods_detailed(sm["id"])

        # Групи для типів періодів
        type_labels = {
            "100": "Періоди 100%",
            "30": "Періоди 30%",
            "non_involved": "Періоди не залучення"
        }

        for period_type, type_label in type_labels.items():
            period_list = periods.get(period_type, [])

            group = QGroupBox(f"{type_label} ({len(period_list)})")
            group_layout = QVBoxLayout()

            if period_list:
                for period in period_list:
                    period_widget = self.create_period_widget(period, period_type)
                    group_layout.addWidget(period_widget)
            else:
                empty_label = QLabel("Немає періодів")
                empty_label.setStyleSheet("color: gray; font-style: italic;")
                group_layout.addWidget(empty_label)

            group.setLayout(group_layout)
            self.periods_layout.addWidget(group)

        # Додаємо stretch в кінці
        self.periods_layout.addStretch()

    def create_period_widget(self, period, period_type):
        """Створює віджет для одного періоду з кнопками редагування/видалення"""
        widget = QFrame()
        widget.setFrameShape(QFrame.StyledPanel)
        widget.setStyleSheet("QFrame { background-color: #f5f5f5; border-radius: 4px; padding: 4px; }")

        layout = QHBoxLayout()
        layout.setContentsMargins(8, 4, 8, 4)

        # Текст періоду
        period_text = f"з {period['start_date']} по {period['end_date']}"
        label = QLabel(period_text)
        label.setStyleSheet("font-weight: bold;")
        layout.addWidget(label)

        layout.addStretch()

        # Кнопка редагування
        edit_btn = QPushButton("Редагувати")
        edit_btn.setStyleSheet("padding: 4px 8px;")
        edit_btn.clicked.connect(lambda checked, p=period: self.edit_period(p))
        layout.addWidget(edit_btn)

        # Кнопка видалення
        delete_btn = QPushButton("Видалити")
        delete_btn.setStyleSheet("background-color: #f44336; color: white; padding: 4px 8px;")
        delete_btn.clicked.connect(lambda checked, p=period: self.delete_period(p))
        layout.addWidget(delete_btn)

        widget.setLayout(layout)
        return widget

    def edit_period(self, period):
        """Редагує період"""
        dialog = EditPeriodDialog(period, self)
        if dialog.exec() == QDialog.Accepted:
            new_start = dialog.start_input.text().strip()
            new_end = dialog.end_input.text().strip()

            # Перевіряємо формат
            try:
                datetime.strptime(new_start, "%d.%m.%Y")
                datetime.strptime(new_end, "%d.%m.%Y")
            except ValueError:
                QMessageBox.warning(self, "Помилка", "Невірний формат дати!")
                return

            try:
                self.db_manager.update_period(period["id"], new_start, new_end)
                QMessageBox.information(self, "Успіх", "Період оновлено!")
                self.load_periods()
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Не вдалось оновити період:\n{str(e)}")

    def delete_period(self, period):
        """Видаляє період"""
        reply = QMessageBox.question(
            self,
            "Підтвердження",
            f"Видалити період з {period['start_date']} по {period['end_date']}?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
                self.db_manager.delete_period(period["id"])
                QMessageBox.information(self, "Успіх", "Період видалено!")
                self.load_periods()
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Не вдалось видалити період:\n{str(e)}")

    def add_new_period(self):
        """Відкриває діалог додавання нового періоду"""
        name = self.name_combo.currentText().strip()
        if not name:
            QMessageBox.warning(self, "Помилка", "Спочатку оберіть військовослужбовця!")
            return

        from gui.add_period_dialog import AddPeriodDialog
        dialog = AddPeriodDialog(self.db_manager, self, preselected_name=name)
        if dialog.exec() == QDialog.Accepted:
            self.load_periods()


class EditPeriodDialog(QDialog):
    """Маленький діалог для редагування дат періоду"""

    def __init__(self, period, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редагування періоду")
        self.setMinimumWidth(300)

        layout = QVBoxLayout()

        # Стиль для полів введення
        input_style = "QLineEdit { color: black; background-color: white; }"

        # Дата початку
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Початок:"))
        self.start_input = QLineEdit(period["start_date"])
        self.start_input.setStyleSheet(input_style)
        start_layout.addWidget(self.start_input)
        layout.addLayout(start_layout)

        # Дата кінця
        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("Кінець:"))
        self.end_input = QLineEdit(period["end_date"])
        self.end_input.setStyleSheet(input_style)
        end_layout.addWidget(self.end_input)
        layout.addLayout(end_layout)

        # Кнопки
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Зберегти")
        save_btn.clicked.connect(self.accept)
        button_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Скасувати")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)
