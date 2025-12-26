"""
Діалог для додавання нового періоду
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QMessageBox,
    QGroupBox, QCompleter
)
from PySide6.QtCore import Qt
from datetime import datetime
import calendar


class AddPeriodDialog(QDialog):
    """
    Діалог для додавання нового періоду для військовослужбовця
    """

    def __init__(self, db_manager, parent=None, preselected_name=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.preselected_name = preselected_name
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Додати період")
        self.setMinimumWidth(500)

        layout = QVBoxLayout()

        # Вибір військовослужбовця
        sm_group = QGroupBox("Військовослужбовець")
        sm_layout = QVBoxLayout()

        # Стилі для полів введення
        combo_style = "QComboBox { color: black; background-color: white; }"
        input_style = "QLineEdit { color: black; background-color: white; }"
        popup_style = "QListView { color: black; background-color: white; }"

        # Dropdown з автодоповненням
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("ПІБ:"))
        self.name_combo = QComboBox()
        self.name_combo.setEditable(True)
        self.name_combo.setInsertPolicy(QComboBox.NoInsert)
        self.name_combo.setStyleSheet(combo_style)

        # Заповнюємо список з БД
        try:
            names = self.db_manager.get_unique_names()
            self.name_combo.addItems(names)

            # Автодоповнення
            completer = QCompleter(names)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            completer.setFilterMode(Qt.MatchContains)
            completer.popup().setStyleSheet(popup_style)
            self.name_combo.setCompleter(completer)

            # Якщо є preselected_name - вибираємо його
            if self.preselected_name:
                index = self.name_combo.findText(self.preselected_name)
                if index >= 0:
                    self.name_combo.setCurrentIndex(index)
        except:
            pass

        self.name_combo.currentTextChanged.connect(self.on_name_changed)
        name_layout.addWidget(self.name_combo)
        sm_layout.addLayout(name_layout)

        # Інформація про в/с (автозаповнення)
        info_layout = QHBoxLayout()
        self.rank_label = QLabel("Звання: -")
        self.position_label = QLabel("Посада: -")
        info_layout.addWidget(self.rank_label)
        info_layout.addWidget(self.position_label)
        sm_layout.addLayout(info_layout)

        self.unit_label = QLabel("Підрозділ: -")
        sm_layout.addWidget(self.unit_label)

        sm_group.setLayout(sm_layout)
        layout.addWidget(sm_group)

        # Новий період
        period_group = QGroupBox("Новий період")
        period_layout = QVBoxLayout()

        # Місяць
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Місяць:"))
        self.month_combo = QComboBox()
        self.month_combo.setEditable(True)
        self.month_combo.setStyleSheet(combo_style)

        # Заповнюємо місяцями з БД + поточний/наступний
        try:
            months = self.db_manager.get_available_months()
            # Додаємо поточний місяць якщо його немає
            current_month = datetime.now().strftime("%Y-%m")
            if current_month not in months:
                months.insert(0, current_month)
            self.month_combo.addItems(months)
        except:
            # Якщо БД порожня - додаємо поточний місяць
            current_month = datetime.now().strftime("%Y-%m")
            self.month_combo.addItem(current_month)

        self.month_combo.currentTextChanged.connect(self.on_month_changed)
        month_layout.addWidget(self.month_combo)
        period_layout.addLayout(month_layout)

        # Тип періоду
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Тип періоду:"))
        self.period_type_combo = QComboBox()
        self.period_type_combo.addItems(["100%", "30%", "Не залучення"])
        self.period_type_combo.setStyleSheet(combo_style)
        type_layout.addWidget(self.period_type_combo)
        period_layout.addLayout(type_layout)

        # Дата початку
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Дата початку:"))
        self.start_date_input = QLineEdit()
        self.start_date_input.setPlaceholderText("01.08.2025")
        self.start_date_input.setStyleSheet(input_style)
        start_layout.addWidget(self.start_date_input)
        period_layout.addLayout(start_layout)

        # Дата кінця
        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("Дата кінця:"))
        self.end_date_input = QLineEdit()
        self.end_date_input.setPlaceholderText("31.08.2025")
        self.end_date_input.setStyleSheet(input_style)
        end_layout.addWidget(self.end_date_input)
        period_layout.addLayout(end_layout)

        period_group.setLayout(period_layout)
        layout.addWidget(period_group)

        # Кнопки
        button_layout = QHBoxLayout()

        add_btn = QPushButton("Додати")
        add_btn.setStyleSheet("background-color: #4caf50; color: white; font-weight: bold; padding: 8px 16px;")
        add_btn.clicked.connect(self.add_period)
        button_layout.addWidget(add_btn)

        cancel_btn = QPushButton("Скасувати")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

        # Ініціалізуємо інформацію
        self.on_name_changed()
        self.on_month_changed()

    def on_name_changed(self):
        """Оновлює інформацію про військовослужбовця"""
        name = self.name_combo.currentText().strip()
        if not name:
            self.rank_label.setText("Звання: -")
            self.position_label.setText("Посада: -")
            self.unit_label.setText("Підрозділ: -")
            return

        try:
            sm = self.db_manager.get_servicemember_by_name(name)
            if sm:
                self.rank_label.setText(f"Звання: {sm.get('rank', '-') or '-'}")
                self.position_label.setText(f"Посада: {sm.get('position', '-') or '-'}")
                self.unit_label.setText(f"Підрозділ: {sm.get('unit', '-') or '-'}")
            else:
                self.rank_label.setText("Звання: (не знайдено)")
                self.position_label.setText("Посада: -")
                self.unit_label.setText("Підрозділ: -")
        except:
            pass

    def on_month_changed(self):
        """Автозаповнення дат при зміні місяця"""
        month_str = self.month_combo.currentText().strip()

        # Очікуємо формат YYYY-MM
        if len(month_str) >= 7 and '-' in month_str:
            try:
                parts = month_str.split('-')
                year = int(parts[0])
                month = int(parts[1])

                # Визначаємо останній день місяця
                last_day = calendar.monthrange(year, month)[1]

                # Форматуємо дати
                start_date = f"01.{month:02d}.{year}"
                end_date = f"{last_day}.{month:02d}.{year}"

                self.start_date_input.setText(start_date)
                self.end_date_input.setText(end_date)
            except:
                pass

    def add_period(self):
        """Додає період для військовослужбовця"""
        # Перевірка
        name = self.name_combo.currentText().strip()
        if not name:
            QMessageBox.warning(self, "Помилка", "Оберіть військовослужбовця!")
            return

        month = self.month_combo.currentText().strip()
        if not month:
            QMessageBox.warning(self, "Помилка", "Вкажіть місяць!")
            return

        start_date = self.start_date_input.text().strip()
        end_date = self.end_date_input.text().strip()

        if not start_date or not end_date:
            QMessageBox.warning(self, "Помилка", "Вкажіть дати початку та кінця періоду!")
            return

        # Перевіряємо формат дат
        try:
            datetime.strptime(start_date, "%d.%m.%Y")
            datetime.strptime(end_date, "%d.%m.%Y")
        except ValueError:
            QMessageBox.warning(self, "Помилка", "Невірний формат дати! Використовуйте DD.MM.YYYY")
            return

        # Отримуємо ID військовослужбовця
        sm = self.db_manager.get_servicemember_by_name(name)
        if not sm:
            QMessageBox.warning(self, "Помилка", f"Військовослужбовця '{name}' не знайдено в базі!")
            return

        # Визначаємо тип періоду
        period_type_text = self.period_type_combo.currentText()
        if period_type_text == "100%":
            period_type = "100"
        elif period_type_text == "30%":
            period_type = "30"
        else:
            period_type = "non_involved"

        try:
            # Додаємо період
            self.db_manager.add_single_period(
                servicemember_id=sm["id"],
                month=month,
                period_type=period_type,
                start_date=start_date,
                end_date=end_date
            )

            QMessageBox.information(
                self,
                "Успіх",
                f"Період {period_type_text} додано для '{name}'!\n"
                f"з {start_date} по {end_date}"
            )

            self.accept()

        except Exception as e:
            QMessageBox.critical(
                self,
                "Помилка",
                f"Не вдалось додати період:\n{str(e)}"
            )
