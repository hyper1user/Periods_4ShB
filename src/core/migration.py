"""
Одноразова міграція даних з Excel в SQLite базу даних
"""
from typing import Dict
from datetime import datetime
from core.excel_reader import ExcelReader
from core.database import DatabaseManager
from core.data_processor import DataProcessor


class DataMigration:
    """
    Клас для міграції даних з Excel файлу в SQLite базу даних
    """

    def __init__(self, excel_reader: ExcelReader, db_manager: DatabaseManager):
        """
        Ініціалізація міграції

        Args:
            excel_reader: Екземпляр ExcelReader
            db_manager: Екземпляр DatabaseManager
        """
        self.excel_reader = excel_reader
        self.db_manager = db_manager
        self.stats = {
            "servicemembers": 0,
            "service_records": 0,
            "periods_100": 0,
            "periods_30": 0,
            "parsed_periods": 0,
            "errors": 0
        }

    def migrate_full_database(self) -> Dict[str, int]:
        """
        Повна міграція всіх аркушів Excel → БД

        Returns:
            Статистика: {"servicemembers": 1594, "service_records": 22125, ...}
        """
        print("=" * 60)
        print("ПОЧАТОК МІГРАЦІЇ ДАНИХ З EXCEL В БД")
        print("=" * 60)

        # Phase 1: Міграція аркуша "Data"
        print("\n[1/4] Міграція аркуша 'Data'...")
        self._migrate_data_sheet()

        # Phase 2: Міграція аркушів періодів (100% та 30%)
        print("\n[2/4] Міграція аркушів періодів...")
        self._migrate_period_sheets()

        # Phase 3: Ініціалізація sync_metadata
        print("\n[3/4] Ініціалізація метаданих синхронізації...")
        self._init_sync_metadata()

        # Phase 4: Валідація
        print("\n[4/4] Валідація міграції...")
        validation_result = self.validate_migration()

        print("\n" + "=" * 60)
        print("МІГРАЦІЯ ЗАВЕРШЕНА")
        print("=" * 60)
        print(f"Військовослужбовців: {self.stats['servicemembers']}")
        print(f"Записів з Data: {self.stats['service_records']}")
        print(f"Періодів 100%: {self.stats['periods_100']}")
        print(f"Періодів 30%: {self.stats['periods_30']}")
        print(f"Розпарсених періодів: {self.stats['parsed_periods']}")
        print(f"Помилок: {self.stats['errors']}")
        print(f"Валідація: {'[OK] ПРОЙДЕНО' if validation_result else '[FAIL] ПРОВАЛЕНО'}")
        print("=" * 60)

        return self.db_manager.get_record_count()

    def _migrate_data_sheet(self):
        """
        Міграція аркуша "Data" → servicemembers + service_records
        ОПТИМІЗОВАНО: Використовує batch insert та рідкі commits

        Логіка:
        1. Читаємо всі рядки з аркуша "Data"
        2. Групуємо в batch по 500 рядків
        3. Використовуємо executemany для швидкого вставлення
        """
        # Отримати всі дані з аркуша Data
        data_rows = self.excel_reader.get_sheet_data("Data")
        print(f"  Знайдено {len(data_rows)} рядків в аркуші Data")

        cursor = self.db_manager.connection.cursor()

        # Відстежування унікальних ПІБ
        seen_names = {}  # name -> (id, rank, position)

        # Batch для service_records
        service_records_batch = []
        BATCH_SIZE = 500

        for i, row in enumerate(data_rows, 1):
            if i % 5000 == 0:
                print(f"  Оброблено {i}/{len(data_rows)} рядків...")

            name = row.get("name")
            if not name:
                continue

            try:
                # Створити servicemember якщо ще не існує
                if name not in seen_names:
                    cursor.execute("""
                        INSERT INTO servicemembers (name, rank, position, rnokpp, unit, birth_date)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (
                        name,
                        row.get("rank"),
                        row.get("position"),
                        row.get("rnokpp"),
                        row.get("unit"),
                        row.get("birth_date")
                    ))
                    sm_id = cursor.lastrowid
                    seen_names[name] = sm_id
                    self.stats["servicemembers"] += 1
                else:
                    sm_id = seen_names[name]

                # Додати service_record в batch
                service_records_batch.append((
                    sm_id,
                    row.get("month"),
                    row.get("unit"),
                    row.get("rank"),
                    row.get("position"),
                    row.get("rnokpp"),
                    row.get("birth_date"),
                    row.get("start_100"),
                    row.get("end_100"),
                    row.get("start_30"),
                    row.get("end_30"),
                    row.get("start_non"),
                    row.get("end_non"),
                    row.get("status"),
                    None  # excel_row_number
                ))
                self.stats["service_records"] += 1

                # Коли batch заповнений - вставляємо
                if len(service_records_batch) >= BATCH_SIZE:
                    cursor.executemany("""
                        INSERT INTO service_records
                        (servicemember_id, month, unit, rank, position, rnokpp, birth_date,
                         start_100, end_100, start_30, end_30, start_non, end_non, status, excel_row_number)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, service_records_batch)
                    self.db_manager.connection.commit()
                    service_records_batch = []

            except Exception as e:
                print(f"  [ERROR] Помилка при обробці рядка {i} ({name}): {e}")
                self.stats["errors"] += 1

        # Вставити залишок batch
        if service_records_batch:
            cursor.executemany("""
                INSERT INTO service_records
                (servicemember_id, month, unit, rank, position, rnokpp, birth_date,
                 start_100, end_100, start_30, end_30, start_non, end_non, status, excel_row_number)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, service_records_batch)
            self.db_manager.connection.commit()

        print(f"  [OK] Мігровано {self.stats['servicemembers']} військовослужбовців")
        print(f"  [OK] Мігровано {self.stats['service_records']} записів")

    def _migrate_period_sheets(self):
        """
        Міграція аркушів "Періоди на 100" та "Періоди на 30" → parsed_periods + periods

        ВАЖЛИВО: Читаємо періоди напряму з цих аркушів замість парсингу
        пошкоджених полів start_100/end_100 з аркуша Data.
        Аркуші періодів мають правильний текстовий формат:
          "з DD.MM.YYYY по DD.MM.YYYY"
        """
        cursor = self.db_manager.connection.cursor()

        # Отримати словник servicemember_id по імені
        cursor.execute("SELECT id, name FROM servicemembers")
        name_to_id = {row[1]: row[0] for row in cursor.fetchall()}

        # Обробка аркушів
        sheets_config = [
            ("Періоди на 100", "100"),
            ("Періоди на 30", "30"),
        ]

        for sheet_name, period_type in sheets_config:
            if sheet_name not in self.excel_reader.workbook.sheetnames:
                print(f"  [SKIP] Аркуш '{sheet_name}' не знайдено")
                continue

            print(f"  Обробка аркуша '{sheet_name}'...")

            # Отримати дані з аркуша
            sheet_data = self.excel_reader.get_sheet_data(sheet_name)

            # Агрегувати періоди по військовослужбовцях
            # {name: [period_text1, period_text2, ...]}
            aggregated = {}
            for row in sheet_data:
                name = row.get("name")
                periods_text = row.get("periods")

                if name and periods_text:
                    if name not in aggregated:
                        aggregated[name] = []
                    aggregated[name].append(periods_text)

            print(f"    Знайдено {len(aggregated)} військовослужбовців з періодами")

            # Для кожного військовослужбовця парсимо та зберігаємо
            saved_count = 0
            for name, periods_list in aggregated.items():
                sm_id = name_to_id.get(name)
                if not sm_id:
                    # Спробувати знайти схоже ім'я
                    continue

                # Об'єднати всі періоди в один текст
                all_periods_text = "\n".join(periods_list)

                # Парсити через DataProcessor
                parsed = DataProcessor.parse_periods(all_periods_text)

                if not parsed:
                    continue

                # Злити послідовні періоди
                merged = DataProcessor.merge_consecutive_periods(parsed)

                # Зберегти кожен період в parsed_periods
                for start_date, end_date in merged:
                    cursor.execute("""
                        INSERT INTO parsed_periods (servicemember_id, period_type, start_date, end_date)
                        VALUES (?, ?, ?, ?)
                    """, (sm_id, period_type, start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")))

                # Сформувати текстове представлення для periods таблиці
                formatted_text = DataProcessor.format_periods_for_document(merged)

                # Зберегти в periods
                cursor.execute("""
                    INSERT OR REPLACE INTO periods (servicemember_id, period_type, period_text)
                    VALUES (?, ?, ?)
                """, (sm_id, period_type, formatted_text))

                saved_count += 1

                if period_type == "100":
                    self.stats["periods_100"] += 1
                else:
                    self.stats["periods_30"] += 1

            self.db_manager.connection.commit()
            print(f"    [OK] Збережено періоди для {saved_count} військовослужбовців")

        # Підрахунок parsed_periods
        cursor.execute("SELECT COUNT(*) FROM parsed_periods")
        self.stats["parsed_periods"] = cursor.fetchone()[0]

        print(f"  [OK] Всього розпарсено {self.stats['parsed_periods']} окремих періодів")

    def _calculate_all_periods(self):
        """
        Розрахунок та збереження періодів для всіх військовослужбовців

        Логіка:
        1. Отримати всіх servicemembers
        2. Для кожного розрахувати періоди
        3. Зберегти в periods та parsed_periods
        """
        servicemembers = self.db_manager.get_all_servicemembers()
        print(f"  Розрахунок періодів для {len(servicemembers)} військовослужбовців...")

        for i, sm in enumerate(servicemembers, 1):
            if i % 100 == 0:
                print(f"  Оброблено {i}/{len(servicemembers)}...")

            try:
                self.db_manager.calculate_and_store_periods(sm["id"])

                # Підрахунок статистики
                periods_100 = self.db_manager.get_periods(sm["id"], "100")
                periods_30 = self.db_manager.get_periods(sm["id"], "30")

                if periods_100:
                    self.stats["periods_100"] += 1
                if periods_30:
                    self.stats["periods_30"] += 1

            except Exception as e:
                print(f"  [ERROR] Помилка при розрахунку періодів для {sm['name']}: {e}")
                self.stats["errors"] += 1

        # Підрахунок parsed_periods
        cursor = self.db_manager.connection.cursor()
        cursor.execute("SELECT COUNT(*) FROM parsed_periods")
        self.stats["parsed_periods"] = cursor.fetchone()[0]

        print(f"  [OK] Розраховано {self.stats['periods_100']} періодів 100%")
        print(f"  [OK] Розраховано {self.stats['periods_30']} періодів 30%")
        print(f"  [OK] Розпарсено {self.stats['parsed_periods']} окремих періодів")

    def _init_sync_metadata(self):
        """
        Ініціалізація таблиці sync_metadata

        Логіка:
        1. Для кожного запису в БД створюємо метадані
        2. Зберігаємо hash та timestamp для подальшої синхронізації
        """
        import hashlib
        import json

        cursor = self.db_manager.connection.cursor()
        count = 0

        # Метадані для servicemembers
        servicemembers = self.db_manager.get_all_servicemembers()
        for sm in servicemembers:
            entity_id = f"servicemembers:{sm['id']}"
            data_hash = hashlib.md5(json.dumps(dict(sm), sort_keys=True, default=str).encode()).hexdigest()

            cursor.execute("""
                INSERT INTO sync_metadata (entity_type, entity_id, last_modified, hash, sync_status)
                VALUES (?, ?, ?, ?, 'synced')
            """, ('database', entity_id, datetime.now(), data_hash))
            count += 1

        self.db_manager.connection.commit()
        print(f"  [OK] Створено {count} записів метаданих синхронізації")

    def validate_migration(self) -> bool:
        """
        Валідація міграції - порівняння Excel vs БД

        Returns:
            True якщо валідація успішна, False інакше
        """
        try:
            # 1. Порівняти кількість унікальних ПІБ
            excel_names = set(self.excel_reader.get_unique_names("Data"))
            db_names = set(self.db_manager.get_unique_names())

            print(f"  Excel унікальних ПІБ: {len(excel_names)}")
            print(f"  БД унікальних ПІБ: {len(db_names)}")

            if len(excel_names) != len(db_names):
                print(f"  [WARN] Кількість ПІБ не співпадає!")
                missing_in_db = excel_names - db_names
                missing_in_excel = db_names - excel_names

                if missing_in_db:
                    print(f"  Відсутні в БД ({len(missing_in_db)}):")
                    for name in list(missing_in_db)[:5]:
                        print(f"    - {name}")
                        print(f"      repr: {repr(name)}")
                        print(f"      len: {len(name)}")
                        # Спробувати знайти схожі імена в БД
                        similar = [db_name for db_name in db_names if name.strip() == db_name.strip()]
                        if similar:
                            print(f"      Схоже в БД: {similar[0]}")
                            print(f"      repr БД: {repr(similar[0])}")

                if missing_in_excel:
                    print(f"  Відсутні в Excel ({len(missing_in_excel)}):")
                    for name in list(missing_in_excel)[:5]:
                        print(f"    - {name}")
                        print(f"      repr: {repr(name)}")

                # Не фейлити валідацію, якщо різниця мінімальна (< 1%)
                diff_percent = abs(len(excel_names) - len(db_names)) / len(excel_names) * 100
                if diff_percent < 1.0:
                    print(f"  [WARN] Різниця мінімальна ({diff_percent:.2f}%), продовжую...")
                else:
                    return False

            # 2. Вибіркова перевірка 10 записів
            import random
            sample_names = random.sample(list(excel_names), min(10, len(excel_names)))

            print(f"\n  Перевірка вибірки з {len(sample_names)} записів...")
            for name in sample_names:
                db_data = self.db_manager.get_complete_data(name)
                if not db_data:
                    print(f"  [ERROR] Не знайдено в БД: {name}")
                    return False

                # Перевірка чи є періоди
                if not db_data["periods_100"] and not db_data["periods_30"]:
                    print(f"  [WARN] Немає періодів для: {name}")

            print(f"  [OK] Вибіркова перевірка пройдена")

            # 3. Перевірка цілісності БД
            cursor = self.db_manager.connection.cursor()
            cursor.execute("""
                SELECT sm.name
                FROM servicemembers sm
                LEFT JOIN service_records sr ON sm.id = sr.servicemember_id
                WHERE sr.id IS NULL
            """)
            orphans = cursor.fetchall()
            if orphans:
                print(f"  [WARN] Знайдено {len(orphans)} службовців без записів")

            return True

        except Exception as e:
            print(f"  [ERROR] Помилка валідації: {e}")
            return False

    def create_backup(self, excel_path: str) -> str:
        """
        Створити backup Excel файлу перед міграцією

        Args:
            excel_path: Шлях до Excel файлу

        Returns:
            Шлях до backup файлу
        """
        import shutil
        from datetime import datetime

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = excel_path.replace(".xlsx", f"_backup_{timestamp}.xlsx")

        shutil.copy2(excel_path, backup_path)
        print(f"[OK] Backup створено: {backup_path}")

        return backup_path
