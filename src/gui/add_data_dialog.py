"""
Діалог для додавання нових даних та періодів
"""
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QTabWidget,
    QWidget, QMessageBox, QDateEdit
)
from PySide6.QtCore import Qt, QDate
from datetime import datetime


class AddDataDialog(QDialog):
    """
    Діалог для додавання даних військовослужбовця та періодів
    """

    def __init__(self, excel_reader, parent=None):
        super().__init__(parent)
        self.excel_reader = excel_reader
        self.init_ui()

    def init_ui(self):
        """
        Ініціалізація інтерфейсу
        """
        self.setWindowTitle("Додати дані")
        self.setMinimumWidth(500)

        layout = QVBoxLayout()

        # Вкладки
        tabs = QTabWidget()

        # Вкладка 1: Додати військовослужбовця
        tab1 = QWidget()
        tab1_layout = QVBoxLayout()

        # Поля для введення
        self.name_input = self.create_field(tab1_layout, "ПІБ (Прізвище Ім'я По-батькові):")
        self.rank_input = self.create_field(tab1_layout, "Звання:")
        self.position_input = self.create_field(tab1_layout, "Посада:")
        self.rnokpp_input = self.create_field(tab1_layout, "РНОКПП:")
        self.unit_input = self.create_field(tab1_layout, "Підрозділ:")

        # Дата народження
        birth_layout = QHBoxLayout()
        birth_layout.addWidget(QLabel("Дата народження (DD.MM.YYYY):"))
        self.birth_date_input = QLineEdit()
        self.birth_date_input.setPlaceholderText("01.01.1990")
        birth_layout.addWidget(self.birth_date_input)
        tab1_layout.addLayout(birth_layout)

        # Місяць
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Місяць (опціонально):"))
        self.month_input = QLineEdit()
        self.month_input.setPlaceholderText("Січень 2024")
        month_layout.addWidget(self.month_input)
        tab1_layout.addLayout(month_layout)

        # Кнопка додати
        self.add_person_btn = QPushButton("Додати військовослужбовця")
        self.add_person_btn.clicked.connect(self.add_servicemember)
        tab1_layout.addWidget(self.add_person_btn)

        tab1_layout.addStretch()
        tab1.setLayout(tab1_layout)

        # Вкладка 2: Додати період
        tab2 = QWidget()
        tab2_layout = QVBoxLayout()

        # Вибір ПІБ
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("ПІБ:"))
        self.period_name_combo = QComboBox()
        self.period_name_combo.setEditable(True)
        try:
            names = self.excel_reader.get_unique_names("Data")
            self.period_name_combo.addItems(names)
        except:
            pass
        name_layout.addWidget(self.period_name_combo)
        tab2_layout.addLayout(name_layout)

        # Тип періоду
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Тип періоду:"))
        self.period_type_combo = QComboBox()
        self.period_type_combo.addItems(["Періоди на 100", "Періоди на 30"])
        type_layout.addWidget(self.period_type_combo)
        tab2_layout.addLayout(type_layout)

        # Дати періоду
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Початок (DD.MM.YYYY):"))
        self.start_date_input = QLineEdit()
        self.start_date_input.setPlaceholderText("01.12.2024")
        start_layout.addWidget(self.start_date_input)
        tab2_layout.addLayout(start_layout)

        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("Кінець (DD.MM.YYYY):"))
        self.end_date_input = QLineEdit()
        self.end_date_input.setPlaceholderText("31.12.2024")
        end_layout.addWidget(self.end_date_input)
        tab2_layout.addLayout(end_layout)

        # Кнопка додати період
        self.add_period_btn = QPushButton("Додати період")
        self.add_period_btn.clicked.connect(self.add_period)
        tab2_layout.addWidget(self.add_period_btn)

        tab2_layout.addStretch()
        tab2.setLayout(tab2_layout)

        # Вкладка 3: Масове додавання періодів
        tab3 = QWidget()
        tab3_layout = QVBoxLayout()

        # Інструкція
        info_label = QLabel("Додати період одразу для багатьох військовослужбовців")
        info_label.setStyleSheet("font-weight: bold; color: #2e7d32;")
        tab3_layout.addWidget(info_label)

        # Вибір режиму
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("Вибрати:"))
        self.mass_mode_combo = QComboBox()
        self.mass_mode_combo.addItems(["Всіх", "За підрозділом", "Вручну"])
        self.mass_mode_combo.currentIndexChanged.connect(self.on_mass_mode_changed)
        mode_layout.addWidget(self.mass_mode_combo)
        tab3_layout.addLayout(mode_layout)

        # Вибір підрозділу (спочатку прихований)
        self.unit_layout = QHBoxLayout()
        self.unit_layout.addWidget(QLabel("Підрозділ:"))
        self.mass_unit_combo = QComboBox()
        try:
            units = self.excel_reader.get_unique_units()
            self.mass_unit_combo.addItems(units)
        except:
            pass
        self.mass_unit_combo.currentIndexChanged.connect(self.update_mass_counter)
        self.unit_layout.addWidget(self.mass_unit_combo)
        self.unit_widget = QWidget()
        self.unit_widget.setLayout(self.unit_layout)
        self.unit_widget.setVisible(False)
        tab3_layout.addWidget(self.unit_widget)

        # Список для вибору вручну (спочатку прихований)
        from PySide6.QtWidgets import QListWidget
        self.manual_list_label = QLabel("Виберіть людей (Ctrl+клік для множинного вибору):")
        self.manual_list_label.setVisible(False)
        tab3_layout.addWidget(self.manual_list_label)

        self.mass_names_list = QListWidget()
        self.mass_names_list.setSelectionMode(QListWidget.MultiSelection)
        try:
            names = self.excel_reader.get_unique_names("Data")
            self.mass_names_list.addItems(names)
        except:
            pass
        self.mass_names_list.itemSelectionChanged.connect(self.update_mass_counter)
        self.mass_names_list.setVisible(False)
        tab3_layout.addWidget(self.mass_names_list)

        # Тип періоду
        mass_type_layout = QHBoxLayout()
        mass_type_layout.addWidget(QLabel("Тип періоду:"))
        self.mass_period_type_combo = QComboBox()
        self.mass_period_type_combo.addItems(["Періоди на 100", "Періоди на 30"])
        mass_type_layout.addWidget(self.mass_period_type_combo)
        tab3_layout.addLayout(mass_type_layout)

        # Місяць
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Місяць (YYYY-MM):"))
        self.mass_month_input = QLineEdit()
        self.mass_month_input.setPlaceholderText("2025-08")
        self.mass_month_input.textChanged.connect(self.on_month_changed)
        month_layout.addWidget(self.mass_month_input)
        tab3_layout.addLayout(month_layout)

        # Дати (автозаповнення)
        mass_start_layout = QHBoxLayout()
        mass_start_layout.addWidget(QLabel("Початок:"))
        self.mass_start_date_input = QLineEdit()
        self.mass_start_date_input.setPlaceholderText("01.08.2025")
        mass_start_layout.addWidget(self.mass_start_date_input)
        tab3_layout.addLayout(mass_start_layout)

        mass_end_layout = QHBoxLayout()
        mass_end_layout.addWidget(QLabel("Кінець:"))
        self.mass_end_date_input = QLineEdit()
        self.mass_end_date_input.setPlaceholderText("31.08.2025")
        mass_end_layout.addWidget(self.mass_end_date_input)
        tab3_layout.addLayout(mass_end_layout)

        # Лічильник
        self.mass_counter_label = QLabel("Буде додано періодів: 0")
        self.mass_counter_label.setStyleSheet("color: #1976d2; font-weight: bold;")
        tab3_layout.addWidget(self.mass_counter_label)

        # Кнопка масового додавання
        self.mass_add_btn = QPushButton("Додати період всім обраним")
        self.mass_add_btn.clicked.connect(self.mass_add_periods)
        self.mass_add_btn.setStyleSheet("background-color: #4caf50; color: white; font-weight: bold; padding: 10px;")
        tab3_layout.addWidget(self.mass_add_btn)

        tab3_layout.addStretch()
        tab3.setLayout(tab3_layout)

        # Додаємо вкладки
        tabs.addTab(tab1, "Додати військовослужбовця")
        tabs.addTab(tab2, "Додати період (один)")
        tabs.addTab(tab3, "Масове додавання періодів")

        layout.addWidget(tabs)

        # Кнопки
        button_layout = QHBoxLayout()

        save_btn = QPushButton("Зберегти зміни")
        save_btn.clicked.connect(self.save_and_close)
        button_layout.addWidget(save_btn)

        close_btn = QPushButton("Закрити")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)

        self.setLayout(layout)

        # Початкове оновлення лічильника
        self.update_mass_counter()

    def create_field(self, layout, label_text):
        """
        Створює поле введення з міткою
        """
        field_layout = QHBoxLayout()
        field_layout.addWidget(QLabel(label_text))
        input_field = QLineEdit()
        field_layout.addWidget(input_field)
        layout.addLayout(field_layout)
        return input_field

    def add_servicemember(self):
        """
        Додає нового військовослужбовця
        """
        # Перевірка обов'язкових полів
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Помилка", "ПІБ є обов'язковим полем!")
            return

        # Збираємо дані
        data = {
            "name": name,
            "rank": self.rank_input.text().strip(),
            "position": self.position_input.text().strip(),
            "rnokpp": self.rnokpp_input.text().strip(),
            "unit": self.unit_input.text().strip(),
            "birth_date": self.birth_date_input.text().strip(),
            "month": self.month_input.text().strip()
        }

        # Додаємо в Excel
        if self.excel_reader.add_servicemember_data(data):
            QMessageBox.information(self, "Успіх", f"Військовослужбовця '{name}' додано!")

            # Очищаємо поля
            self.name_input.clear()
            self.rank_input.clear()
            self.position_input.clear()
            self.rnokpp_input.clear()
            self.unit_input.clear()
            self.birth_date_input.clear()
            self.month_input.clear()

            # Оновлюємо список у вкладці періодів
            try:
                self.period_name_combo.clear()
                names = self.excel_reader.get_unique_names("Data")
                self.period_name_combo.addItems(names)
            except:
                pass
        else:
            QMessageBox.critical(self, "Помилка", "Не вдалось додати військовослужбовця!")

    def add_period(self):
        """
        Додає період для військовослужбовця
        """
        name = self.period_name_combo.currentText().strip()
        if not name:
            QMessageBox.warning(self, "Помилка", "Оберіть ПІБ!")
            return

        start_date = self.start_date_input.text().strip()
        end_date = self.end_date_input.text().strip()

        if not start_date or not end_date:
            QMessageBox.warning(self, "Помилка", "Вкажіть дати початку та кінця періоду!")
            return

        # Формуємо текст періоду
        period_text = f"з {start_date} по {end_date}"

        # Тип періоду
        sheet_name = self.period_type_combo.currentText()

        # Додаємо в Excel
        if self.excel_reader.add_period(name, sheet_name, period_text):
            QMessageBox.information(self, "Успіх", f"Період додано для '{name}'!")

            # Очищаємо дати
            self.start_date_input.clear()
            self.end_date_input.clear()
        else:
            QMessageBox.critical(self, "Помилка", "Не вдалось додати період!")

    def on_mass_mode_changed(self):
        """
        Обробник зміни режиму масового додавання
        """
        mode = self.mass_mode_combo.currentText()

        # Показуємо/ховаємо відповідні поля
        self.unit_widget.setVisible(mode == "За підрозділом")
        self.manual_list_label.setVisible(mode == "Вручну")
        self.mass_names_list.setVisible(mode == "Вручну")

        # Оновлюємо лічильник
        self.update_mass_counter()

    def on_month_changed(self):
        """
        Автозаповнення дат при зміні місяця
        """
        month_str = self.mass_month_input.text().strip()

        # Очікуємо формат YYYY-MM
        if len(month_str) == 7 and month_str[4] == '-':
            try:
                year = int(month_str[:4])
                month = int(month_str[5:7])

                # Визначаємо останній день місяця
                import calendar
                last_day = calendar.monthrange(year, month)[1]

                # Форматуємо дати
                start_date = f"01.{month:02d}.{year}"
                end_date = f"{last_day}.{month:02d}.{year}"

                self.mass_start_date_input.setText(start_date)
                self.mass_end_date_input.setText(end_date)
            except:
                pass

        # Оновлюємо лічильник
        self.update_mass_counter()

    def update_mass_counter(self):
        """
        Оновлює лічильник обраних людей
        """
        mode = self.mass_mode_combo.currentText()
        count = 0

        try:
            if mode == "Всіх":
                names = self.excel_reader.get_unique_names("Data")
                count = len(names)
            elif mode == "За підрозділом":
                unit = self.mass_unit_combo.currentText()
                if unit:
                    unit_data = self.excel_reader.get_unit_data(unit, "Data")
                    unique_names = set([row["name"] for row in unit_data if row.get("name")])
                    count = len(unique_names)
            elif mode == "Вручну":
                count = len(self.mass_names_list.selectedItems())
        except:
            pass

        self.mass_counter_label.setText(f"Буде додано періодів: {count}")

    def mass_add_periods(self):
        """
        Масове додавання періодів
        """
        # Перевірка дат
        start_date = self.mass_start_date_input.text().strip()
        end_date = self.mass_end_date_input.text().strip()

        if not start_date or not end_date:
            QMessageBox.warning(self, "Помилка", "Вкажіть дати початку та кінця періоду!")
            return

        # Формуємо текст періоду
        period_text = f"з {start_date} по {end_date}"

        # Тип періоду
        sheet_name = self.mass_period_type_combo.currentText()

        # Визначаємо список людей
        mode = self.mass_mode_combo.currentText()
        names = []

        try:
            if mode == "Всіх":
                names = self.excel_reader.get_unique_names("Data")
            elif mode == "За підрозділом":
                unit = self.mass_unit_combo.currentText()
                if not unit:
                    QMessageBox.warning(self, "Помилка", "Оберіть підрозділ!")
                    return
                unit_data = self.excel_reader.get_unit_data(unit, "Data")
                names = list(set([row["name"] for row in unit_data if row.get("name")]))
            elif mode == "Вручну":
                selected_items = self.mass_names_list.selectedItems()
                if not selected_items:
                    QMessageBox.warning(self, "Помилка", "Оберіть хоча б одну людину!")
                    return
                names = [item.text() for item in selected_items]
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Помилка при отриманні списку: {str(e)}")
            return

        if not names:
            QMessageBox.warning(self, "Помилка", "Не знайдено жодної людини для додавання!")
            return

        # Підтвердження
        reply = QMessageBox.question(
            self,
            "Підтвердження",
            f"Додати період '{period_text}' для {len(names)} людей у '{sheet_name}'?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # Додаємо періоди
        success_count = 0
        error_count = 0

        from PySide6.QtWidgets import QProgressDialog
        progress = QProgressDialog("Додавання періодів...", "Скасувати", 0, len(names), self)
        progress.setWindowModality(Qt.WindowModal)

        for i, name in enumerate(names):
            progress.setValue(i)
            if progress.wasCanceled():
                break

            if self.excel_reader.add_period(name, sheet_name, period_text):
                success_count += 1
            else:
                error_count += 1

        progress.setValue(len(names))

        # Показуємо результат
        message = f"Додано періодів: {success_count}\nПомилок: {error_count}"
        if error_count == 0:
            QMessageBox.information(self, "Успіх", message)
        else:
            QMessageBox.warning(self, "Завершено з помилками", message)

        # Очищаємо поля
        self.mass_month_input.clear()
        self.mass_start_date_input.clear()
        self.mass_end_date_input.clear()
        self.update_mass_counter()

    def save_and_close(self):
        """
        Зберігає зміни та закриває діалог
        """
        if self.excel_reader.save():
            QMessageBox.information(self, "Успіх", "Зміни збережено у файл!")
            self.accept()
        else:
            QMessageBox.critical(self, "Помилка", "Не вдалось зберегти зміни!")
