"""
Управління SQLite базою даних для військових рапортів
"""
import sqlite3
from typing import List, Dict, Optional, Tuple
from datetime import datetime, date
from contextlib import contextmanager
from utils.date_utils import parse_period_string, format_period
from core.data_processor import DataProcessor
from core.dodatky_reader import get_dodatky_reader


class DatabaseManager:
    """
    Клас для управління SQLite базою даних
    Забезпечує CRUD операції та інтеграцію з Excel через синхронізацію
    """

    def __init__(self, db_path: str):
        """
        Ініціалізація менеджера БД

        Args:
            db_path: Шлях до файлу SQLite БД
        """
        self.db_path = db_path
        self.connection = None

    def connect(self):
        """Підключення до БД та створення таблиць якщо не існують"""
        self.connection = sqlite3.connect(self.db_path)
        self.connection.row_factory = sqlite3.Row  # Доступ через імена колонок
        self._create_tables()
        self._create_triggers()

    def close(self):
        """Закрити з'єднання з БД"""
        if self.connection:
            self.connection.close()
            self.connection = None

    @contextmanager
    def transaction(self):
        """Context manager для транзакцій"""
        try:
            yield self.connection
            self.connection.commit()
        except Exception as e:
            self.connection.rollback()
            raise

    def _create_tables(self):
        """Створення всіх таблиць БД"""
        cursor = self.connection.cursor()

        # Таблиця servicemembers
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS servicemembers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                rank TEXT,
                position TEXT,
                rnokpp TEXT,
                unit TEXT,
                birth_date TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_servicemembers_name
            ON servicemembers(name)
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_servicemembers_unit
            ON servicemembers(unit)
        """)

        # Таблиця service_records
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS service_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                servicemember_id INTEGER NOT NULL,
                month TEXT,
                unit TEXT,
                rank TEXT,
                position TEXT,
                rnokpp TEXT,
                birth_date TEXT,
                start_100 TEXT,
                end_100 TEXT,
                start_30 TEXT,
                end_30 TEXT,
                start_non TEXT,
                end_non TEXT,
                status TEXT,
                excel_row_number INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (servicemember_id) REFERENCES servicemembers(id) ON DELETE CASCADE
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_service_records_member
            ON service_records(servicemember_id)
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_service_records_month
            ON service_records(month)
        """)

        # Таблиця periods
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                servicemember_id INTEGER NOT NULL,
                period_type TEXT CHECK(period_type IN ('100', '30', 'non_involved')),
                period_text TEXT,
                month TEXT,
                excel_row_number INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (servicemember_id) REFERENCES servicemembers(id) ON DELETE CASCADE
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_periods_member_type
            ON periods(servicemember_id, period_type)
        """)

        # Таблиця parsed_periods
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS parsed_periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                servicemember_id INTEGER NOT NULL,
                period_type TEXT CHECK(period_type IN ('100', '30', 'non_involved')),
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                source_record_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (servicemember_id) REFERENCES servicemembers(id) ON DELETE CASCADE
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_parsed_periods_member
            ON parsed_periods(servicemember_id)
        """)

        # Таблиця sync_metadata
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sync_metadata (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_type TEXT NOT NULL,
                entity_id TEXT,
                last_modified TIMESTAMP NOT NULL,
                hash TEXT,
                sync_status TEXT DEFAULT 'synced' CHECK(sync_status IN ('synced', 'conflict', 'pending'))
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_sync_entity
            ON sync_metadata(entity_type, entity_id)
        """)

        # Таблиця sync_log
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sync_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sync_timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                direction TEXT CHECK(direction IN ('excel_to_db', 'db_to_excel', 'bidirectional')),
                records_updated INTEGER,
                conflicts_detected INTEGER,
                status TEXT,
                error_message TEXT
            )
        """)

        # Таблиця ЖБД (журнали бойових дій)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS zbd (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                number TEXT NOT NULL,
                date TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_zbd_date
            ON zbd(date)
        """)

        # Таблиця громад
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS hromady (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_hromady_dates
            ON hromady(start_date, end_date)
        """)

        # Таблиця населених пунктів
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS naseleni_punkty (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_np_dates
            ON naseleni_punkty(start_date, end_date)
        """)

        # View для швидкого доступу
        cursor.execute("""
            CREATE VIEW IF NOT EXISTS v_servicemember_complete AS
            SELECT
                sm.id,
                sm.name,
                sm.rank,
                sm.position,
                sm.rnokpp,
                sm.unit,
                sm.birth_date,
                p100.period_text as periods_100,
                p30.period_text as periods_30
            FROM servicemembers sm
            LEFT JOIN periods p100 ON sm.id = p100.servicemember_id AND p100.period_type = '100'
            LEFT JOIN periods p30 ON sm.id = p30.servicemember_id AND p30.period_type = '30'
        """)

        self.connection.commit()

    def _create_triggers(self):
        """Створення тригерів для автоматичного оновлення"""
        cursor = self.connection.cursor()

        cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS update_periods_on_service_change
            AFTER INSERT ON service_records
            BEGIN
                UPDATE servicemembers
                SET updated_at = CURRENT_TIMESTAMP
                WHERE id = NEW.servicemember_id;
            END
        """)

        cursor.execute("""
            CREATE TRIGGER IF NOT EXISTS update_periods_on_service_update
            AFTER UPDATE ON service_records
            BEGIN
                UPDATE servicemembers
                SET updated_at = CURRENT_TIMESTAMP
                WHERE id = NEW.servicemember_id;
            END
        """)

        self.connection.commit()

    # ==================== CRUD для servicemembers ====================

    def add_servicemember(self, data: Dict) -> int:
        """
        Додати військовослужбовця

        Args:
            data: Словник з даними {name, rank, position, rnokpp, unit, birth_date}

        Returns:
            ID створеного запису
        """
        cursor = self.connection.cursor()

        # Спроба знайти існуючого
        existing = self.get_servicemember_by_name(data["name"])
        if existing:
            # Оновити існуючого
            self.update_servicemember(existing["id"], data)
            return existing["id"]

        cursor.execute("""
            INSERT INTO servicemembers (name, rank, position, rnokpp, unit, birth_date)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            data.get("name"),
            data.get("rank"),
            data.get("position"),
            data.get("rnokpp"),
            data.get("unit"),
            data.get("birth_date")
        ))

        self.connection.commit()
        return cursor.lastrowid

    def get_servicemember_by_name(self, name: str) -> Optional[Dict]:
        """Знайти військовослужбовця по ПІБ"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT * FROM servicemembers WHERE name = ?
        """, (name,))

        row = cursor.fetchone()
        return dict(row) if row else None

    def get_servicemember_by_id(self, id: int) -> Optional[Dict]:
        """Знайти військовослужбовця по ID"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT * FROM servicemembers WHERE id = ?
        """, (id,))

        row = cursor.fetchone()
        return dict(row) if row else None

    def update_servicemember(self, id: int, data: Dict):
        """Оновити дані військовослужбовця"""
        cursor = self.connection.cursor()
        cursor.execute("""
            UPDATE servicemembers
            SET rank = ?, position = ?, rnokpp = ?, unit = ?, birth_date = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (
            data.get("rank"),
            data.get("position"),
            data.get("rnokpp"),
            data.get("unit"),
            data.get("birth_date"),
            id
        ))
        self.connection.commit()

    def get_all_servicemembers(self) -> List[Dict]:
        """Отримати всіх військовослужбовців"""
        cursor = self.connection.cursor()
        cursor.execute("SELECT * FROM servicemembers ORDER BY name")
        return [dict(row) for row in cursor.fetchall()]

    def get_unique_names(self) -> List[str]:
        """Отримати список унікальних ПІБ (аналог excel_reader.get_unique_names)"""
        cursor = self.connection.cursor()
        cursor.execute("SELECT DISTINCT name FROM servicemembers ORDER BY name")
        return [row[0] for row in cursor.fetchall()]

    def get_unique_units(self) -> List[str]:
        """Отримати список унікальних підрозділів (аналог excel_reader.get_unique_units)"""
        cursor = self.connection.cursor()
        cursor.execute("SELECT DISTINCT unit FROM servicemembers WHERE unit IS NOT NULL ORDER BY unit")
        return [row[0] for row in cursor.fetchall()]

    # ==================== CRUD для service_records ====================

    def add_service_record(self, servicemember_id: int, data: Dict) -> int:
        """
        Додати запис з аркуша Data

        Args:
            servicemember_id: ID військовослужбовця
            data: Словник з даними (month, unit, rank, ...)

        Returns:
            ID створеного запису
        """
        cursor = self.connection.cursor()
        cursor.execute("""
            INSERT INTO service_records (
                servicemember_id, month, unit, rank, position, rnokpp, birth_date,
                start_100, end_100, start_30, end_30, start_non, end_non, status, excel_row_number
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            servicemember_id,
            data.get("month"),
            data.get("unit"),
            data.get("rank"),
            data.get("position"),
            data.get("rnokpp"),
            data.get("birth_date"),
            data.get("start_100"),
            data.get("end_100"),
            data.get("start_30"),
            data.get("end_30"),
            data.get("start_non"),
            data.get("end_non"),
            data.get("status"),
            data.get("row_number")
        ))

        self.connection.commit()
        return cursor.lastrowid

    def get_service_records(self, servicemember_id: int) -> List[Dict]:
        """Отримати всі записи для військовослужбовця"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT * FROM service_records WHERE servicemember_id = ? ORDER BY month
        """, (servicemember_id,))
        return [dict(row) for row in cursor.fetchall()]

    # ==================== Робота з періодами ====================

    def calculate_and_store_periods(self, servicemember_id: int):
        """
        Розрахувати та зберегти періоди для військовослужбовця

        Логіка:
        1. Отримати всі service_records
        2. Зібрати та розпарсити періоди 100% та 30%
        3. Злити послідовні періоди
        4. Зберегти в таблиці periods та parsed_periods
        """
        # Отримати всі записи
        records = self.get_service_records(servicemember_id)

        # Збір періодів
        periods_100 = []
        periods_30 = []

        def parse_date(date_str):
            """Парсить дату з різних форматів"""
            if not date_str:
                return None

            date_str = str(date_str).strip()

            # Спробувати різні формати
            formats = [
                "%d.%m.%Y",                    # 01.08.2025
                "%Y-%m-%d",                    # 2025-08-01
                "%Y-%m-%d %H:%M:%S",           # 2025-08-01 00:00:00
                "%Y-%m-%dT%H:%M:%S",           # 2025-08-01T00:00:00
                "%Y-%m-%d %H:%M:%S.%f",        # 2025-08-01 00:00:00.000000
            ]

            for fmt in formats:
                try:
                    return datetime.strptime(date_str, fmt).date()
                except ValueError:
                    continue

            # Якщо нічого не спрацювало - спробувати fromisoformat
            try:
                return datetime.fromisoformat(date_str.replace(" ", "T")).date()
            except:
                return None

        for record in records:
            # Періоди 100%
            if record["start_100"] and record["end_100"]:
                start = parse_date(record["start_100"])
                end = parse_date(record["end_100"])

                if start and end:
                    periods_100.append((start, end))

            # Періоди 30%
            if record["start_30"] and record["end_30"]:
                start = parse_date(record["start_30"])
                end = parse_date(record["end_30"])

                if start and end:
                    periods_30.append((start, end))

        # Злиття послідовних періодів
        merged_100 = DataProcessor.merge_consecutive_periods(periods_100)
        merged_30 = DataProcessor.merge_consecutive_periods(periods_30)

        # Форматування для збереження
        formatted_100 = DataProcessor.format_periods_for_document(merged_100)
        formatted_30 = DataProcessor.format_periods_for_document(merged_30)

        cursor = self.connection.cursor()

        # Видалити старі періоди
        cursor.execute("DELETE FROM periods WHERE servicemember_id = ?", (servicemember_id,))
        cursor.execute("DELETE FROM parsed_periods WHERE servicemember_id = ?", (servicemember_id,))

        # Зберегти нові періоди
        if formatted_100:
            cursor.execute("""
                INSERT INTO periods (servicemember_id, period_type, period_text)
                VALUES (?, '100', ?)
            """, (servicemember_id, formatted_100))

        if formatted_30:
            cursor.execute("""
                INSERT INTO periods (servicemember_id, period_type, period_text)
                VALUES (?, '30', ?)
            """, (servicemember_id, formatted_30))

        # Зберегти parsed_periods
        for start, end in merged_100:
            cursor.execute("""
                INSERT INTO parsed_periods (servicemember_id, period_type, start_date, end_date)
                VALUES (?, '100', ?, ?)
            """, (servicemember_id, start.isoformat(), end.isoformat()))

        for start, end in merged_30:
            cursor.execute("""
                INSERT INTO parsed_periods (servicemember_id, period_type, start_date, end_date)
                VALUES (?, '30', ?, ?)
            """, (servicemember_id, start.isoformat(), end.isoformat()))

        self.connection.commit()

    def get_periods(self, servicemember_id: int, period_type: str) -> str:
        """Отримати текст періодів для військовослужбовця"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT period_text FROM periods
            WHERE servicemember_id = ? AND period_type = ?
        """, (servicemember_id, period_type))

        row = cursor.fetchone()
        return row[0] if row else ""

    # ==================== Головний метод - аналог aggregate_servicemember_data ====================

    def get_complete_data(self, name: str) -> Optional[Dict]:
        """
        Отримати повні дані для військовослужбовця (аналог DataProcessor.aggregate_servicemember_data)

        Returns:
            Словник з даними у форматі DataProcessor:
            {
                "name": "...", "rank": "...", "position": "...", "rnokpp": "...",
                "unit": "...", "birth_date": "...",
                "periods": "...",        # Всі періоди
                "periods_100": "...",    # Тільки 100%
                "periods_30": "...",     # Тільки 30%
                "periods_all": "...",    # Всі разом
                "periods_list": [...]    # Список кортежів
            }
        """
        # Отримати основну інформацію
        servicemember = self.get_servicemember_by_name(name)
        if not servicemember:
            return None

        # ВИПРАВЛЕННЯ: Отримати звання та посаду з ОСТАННЬОГО service_record
        # Аналогічно до Excel Reader - беремо найсвіжіші дані
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT rank, position, rnokpp, unit, birth_date, month
            FROM service_records
            WHERE servicemember_id = ?
            ORDER BY month DESC, id DESC
            LIMIT 1
        """, (servicemember["id"],))

        latest_record = cursor.fetchone()
        if latest_record:
            # Оновлюємо дані з останнього запису
            servicemember["rank"] = latest_record[0] or servicemember["rank"]
            servicemember["position"] = latest_record[1] or servicemember["position"]
            # Також шукаємо РНОКПП та дату народження якщо відсутні
            if not servicemember.get("rnokpp") and latest_record[2]:
                servicemember["rnokpp"] = latest_record[2]
            if not servicemember.get("birth_date") and latest_record[4]:
                servicemember["birth_date"] = latest_record[4]

        # Якщо РНОКПП або дата народження все ще порожні - шукаємо у всіх записах
        if not servicemember.get("rnokpp") or not servicemember.get("birth_date"):
            cursor.execute("""
                SELECT rnokpp, birth_date
                FROM service_records
                WHERE servicemember_id = ?
                  AND (rnokpp IS NOT NULL OR birth_date IS NOT NULL)
                ORDER BY id DESC
            """, (servicemember["id"],))

            for row in cursor.fetchall():
                if not servicemember.get("rnokpp") and row[0]:
                    servicemember["rnokpp"] = row[0]
                if not servicemember.get("birth_date") and row[1]:
                    servicemember["birth_date"] = row[1]
                if servicemember.get("rnokpp") and servicemember.get("birth_date"):
                    break

        # Отримати періоди
        periods_100_text = self.get_periods(servicemember["id"], "100")
        periods_30_text = self.get_periods(servicemember["id"], "30")

        # Отримати parsed_periods для periods_list
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT start_date, end_date FROM parsed_periods
            WHERE servicemember_id = ?
            ORDER BY start_date
        """, (servicemember["id"],))

        periods_list = []
        for row in cursor.fetchall():
            start = datetime.fromisoformat(row[0]).date()
            end = datetime.fromisoformat(row[1]).date()
            periods_list.append((start, end))

        # Об'єднати всі періоди для periods_all
        cursor.execute("""
            SELECT start_date, end_date FROM parsed_periods
            WHERE servicemember_id = ? AND period_type IN ('100', '30')
            ORDER BY start_date
        """, (servicemember["id"],))

        all_periods = []
        for row in cursor.fetchall():
            start = datetime.fromisoformat(row[0]).date()
            end = datetime.fromisoformat(row[1]).date()
            all_periods.append((start, end))

        merged_all = DataProcessor.merge_consecutive_periods(all_periods)
        periods_all_text = DataProcessor.format_periods_for_document(merged_all)

        # Перша літера посади - маленька
        position = servicemember["position"]
        if position and len(position) > 0:
            position = position[0].lower() + position[1:]

        # Отримати ЖБД та Громади з Dodatky.xlsx
        try:
            dodatky = get_dodatky_reader()
            zbd_text = dodatky.get_zbd(periods_all_text)
            hromady_text = dodatky.get_hromady(periods_all_text)
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"[ERROR] Помилка читання Dodatky.xlsx: {e}\n{error_details}")
            # Кидаємо exception вверх щоб побачити помилку в діалозі
            raise Exception(f"Не вдалося прочитати Dodatky.xlsx: {e}")

        return {
            "name": servicemember["name"],
            "rank": servicemember["rank"],
            "position": position,
            "rnokpp": servicemember["rnokpp"],
            "unit": servicemember["unit"],
            "birth_date": servicemember["birth_date"],
            "periods": periods_all_text,  # Для Pilgova: {{ПЕРІОДИ}}
            "periods_100": periods_100_text,  # Для Only100: {{ПЕРІОДИ_100}}
            "periods_30": periods_30_text,
            "periods_all": periods_all_text,
            "periods_list": merged_all,
            "zbd": zbd_text,  # {{ЖБД}}
            "hromady": hromady_text  # {{ГРОМАДА}}
        }

    # ==================== Імпорт даних за місяць ====================

    def import_month_data(self, month: str, data, progress_callback=None) -> Dict[str, int]:
        """
        Імпортує дані за новий місяць

        Args:
            month: Місяць у форматі YYYY-MM
            data: Список записів [{name, rank, position, start_100, end_100, ...}, ...]
                  АБО словник {ПІБ: {rank, position, start_100, end_100, ...}} (старий формат)
            progress_callback: Функція для оновлення прогресу (current, total, message)

        Returns:
            Статистика: {"added": 150, "updated": 0, "errors": 0}
        """
        stats = {"added": 0, "updated": 0, "errors": 0}
        updated_servicemembers = set()

        # Перетворюємо дані в єдиний формат (список записів)
        if isinstance(data, dict):
            # Старий формат: словник {ПІБ: {дані}}
            records = []
            for name, record_data in data.items():
                record = dict(record_data)
                record["name"] = name
                records.append(record)
        else:
            # Новий формат: список записів
            records = data

        total_records = len(records)

        for i, record_data in enumerate(records, 1):
            name = record_data.get("name")
            if not name:
                continue

            # Оновлюємо прогрес
            if progress_callback:
                progress_callback(i, total_records, f"Імпорт запису {i}/{total_records}: {name}")

            try:
                # 1. Знайти або створити servicemember
                servicemember = self.get_servicemember_by_name(name)

                if not servicemember:
                    # Створити нового
                    sm_data = {
                        "name": name,
                        "rank": record_data.get("rank", ""),
                        "position": record_data.get("position", ""),
                        "rnokpp": record_data.get("rnokpp", ""),
                        "unit": record_data.get("unit", ""),
                        "birth_date": record_data.get("birth_date", "")
                    }
                    servicemember_id = self.add_servicemember(sm_data)
                else:
                    servicemember_id = servicemember["id"]
                    # Оновити звання та посаду якщо вони змінились
                    if record_data.get("rank") or record_data.get("position"):
                        update_data = {
                            "rank": record_data.get("rank") or servicemember.get("rank"),
                            "position": record_data.get("position") or servicemember.get("position"),
                            "rnokpp": servicemember.get("rnokpp"),
                            "unit": servicemember.get("unit"),
                            "birth_date": servicemember.get("birth_date")
                        }
                        self.update_servicemember(servicemember_id, update_data)

                # 2. Додати service_record за цей місяць
                service_record = {
                    "month": month,
                    "rank": record_data.get("rank", ""),
                    "position": record_data.get("position", ""),
                    "unit": record_data.get("unit", ""),
                    "rnokpp": record_data.get("rnokpp", ""),
                    "birth_date": record_data.get("birth_date", ""),
                    "start_100": record_data.get("start_100"),
                    "end_100": record_data.get("end_100"),
                    "start_30": record_data.get("start_30"),
                    "end_30": record_data.get("end_30"),
                    "start_non": record_data.get("start_non"),
                    "end_non": record_data.get("end_non"),
                    "status": record_data.get("status", ""),
                    "row_number": None  # Для імпорту не потрібен
                }

                self.add_service_record(servicemember_id, service_record)
                updated_servicemembers.add(servicemember_id)
                stats["added"] += 1

            except Exception as e:
                print(f"[ERROR] Помилка при імпорті {name}: {e}")
                stats["errors"] += 1

        # 3. Перерахувати періоди для всіх оновлених servicemembers
        print(f"\nПерерахунок періодів для {len(updated_servicemembers)} військовослужбовців...")
        updated_list = list(updated_servicemembers)
        total_to_update = len(updated_list)

        for idx, sm_id in enumerate(updated_list, 1):
            if progress_callback:
                progress_callback(
                    total_records + idx,
                    total_records + total_to_update,
                    f"Перерахунок періодів {idx}/{total_to_update}"
                )

            try:
                self.calculate_and_store_periods(sm_id)
            except Exception as e:
                print(f"[WARN] Помилка перерахунку періодів для ID={sm_id}: {e}")

        return stats

    # ==================== Утиліти ====================

    def is_empty(self) -> bool:
        """Перевірка чи БД порожня"""
        cursor = self.connection.cursor()
        cursor.execute("SELECT COUNT(*) FROM servicemembers")
        count = cursor.fetchone()[0]
        return count == 0

    def get_record_count(self) -> Dict[str, int]:
        """
        Статистика кількості записів

        Returns:
            {"servicemembers": 1594, "service_records": 22125, ...}
        """
        cursor = self.connection.cursor()

        counts = {}

        cursor.execute("SELECT COUNT(*) FROM servicemembers")
        counts["servicemembers"] = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM service_records")
        counts["service_records"] = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM periods")
        counts["periods"] = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM parsed_periods")
        counts["parsed_periods"] = cursor.fetchone()[0]

        return counts

    # ==================== CRUD для періодів (нові методи) ====================

    def add_single_period(self, servicemember_id: int, month: str,
                          period_type: str, start_date: str, end_date: str) -> int:
        """
        Додає один період для військовослужбовця

        Args:
            servicemember_id: ID військовослужбовця
            month: Місяць (YYYY-MM)
            period_type: '100', '30', або 'non_involved'
            start_date: Дата початку (DD.MM.YYYY)
            end_date: Дата кінця (DD.MM.YYYY)

        Returns:
            ID створеного service_record
        """
        # Визначаємо поля для запису
        record_data = {
            "month": month,
            "start_100": start_date if period_type == "100" else None,
            "end_100": end_date if period_type == "100" else None,
            "start_30": start_date if period_type == "30" else None,
            "end_30": end_date if period_type == "30" else None,
            "start_non": start_date if period_type == "non_involved" else None,
            "end_non": end_date if period_type == "non_involved" else None,
        }

        # Додаємо service_record
        record_id = self.add_service_record(servicemember_id, record_data)

        # Перераховуємо всі періоди
        self.calculate_and_store_periods(servicemember_id)

        return record_id

    def get_servicemember_periods_detailed(self, servicemember_id: int) -> Dict[str, List[Dict]]:
        """
        Отримати всі періоди для військовослужбовця з деталями для редагування

        Returns:
            {
                "100": [{"id": 1, "start_date": "01.08.2025", "end_date": "31.08.2025"}, ...],
                "30": [...],
                "non_involved": [...]
            }
        """
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT id, period_type, start_date, end_date
            FROM parsed_periods
            WHERE servicemember_id = ?
            ORDER BY period_type, start_date
        """, (servicemember_id,))

        result = {"100": [], "30": [], "non_involved": []}

        for row in cursor.fetchall():
            period_type = row[1]
            # Конвертуємо дату з ISO в DD.MM.YYYY
            start_date = datetime.fromisoformat(row[2]).strftime("%d.%m.%Y")
            end_date = datetime.fromisoformat(row[3]).strftime("%d.%m.%Y")

            if period_type in result:
                result[period_type].append({
                    "id": row[0],
                    "start_date": start_date,
                    "end_date": end_date
                })

        return result

    def update_period(self, period_id: int, start_date: str, end_date: str):
        """
        Оновлює існуючий період

        Args:
            period_id: ID періоду в parsed_periods
            start_date: Нова дата початку (DD.MM.YYYY)
            end_date: Нова дата кінця (DD.MM.YYYY)
        """
        cursor = self.connection.cursor()

        # Конвертуємо дату в ISO формат
        start_iso = datetime.strptime(start_date, "%d.%m.%Y").date().isoformat()
        end_iso = datetime.strptime(end_date, "%d.%m.%Y").date().isoformat()

        # Отримуємо servicemember_id перед оновленням
        cursor.execute("SELECT servicemember_id FROM parsed_periods WHERE id = ?", (period_id,))
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"Період з ID {period_id} не знайдено")

        servicemember_id = row[0]

        # Оновлюємо період
        cursor.execute("""
            UPDATE parsed_periods
            SET start_date = ?, end_date = ?
            WHERE id = ?
        """, (start_iso, end_iso, period_id))

        self.connection.commit()

        # Перераховуємо текстове представлення періодів
        self._recalculate_period_text(servicemember_id)

    def delete_period(self, period_id: int):
        """
        Видаляє період

        Args:
            period_id: ID періоду в parsed_periods
        """
        cursor = self.connection.cursor()

        # Отримуємо servicemember_id перед видаленням
        cursor.execute("SELECT servicemember_id FROM parsed_periods WHERE id = ?", (period_id,))
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"Період з ID {period_id} не знайдено")

        servicemember_id = row[0]

        # Видаляємо період
        cursor.execute("DELETE FROM parsed_periods WHERE id = ?", (period_id,))
        self.connection.commit()

        # Перераховуємо текстове представлення періодів
        self._recalculate_period_text(servicemember_id)

    def _recalculate_period_text(self, servicemember_id: int):
        """
        Перераховує текстове представлення періодів з parsed_periods

        Це потрібно після редагування/видалення періодів напряму в parsed_periods
        """
        cursor = self.connection.cursor()

        # Отримуємо всі періоди 100%
        cursor.execute("""
            SELECT start_date, end_date FROM parsed_periods
            WHERE servicemember_id = ? AND period_type = '100'
            ORDER BY start_date
        """, (servicemember_id,))

        periods_100 = []
        for row in cursor.fetchall():
            start = datetime.fromisoformat(row[0]).date()
            end = datetime.fromisoformat(row[1]).date()
            periods_100.append((start, end))

        # Отримуємо всі періоди 30%
        cursor.execute("""
            SELECT start_date, end_date FROM parsed_periods
            WHERE servicemember_id = ? AND period_type = '30'
            ORDER BY start_date
        """, (servicemember_id,))

        periods_30 = []
        for row in cursor.fetchall():
            start = datetime.fromisoformat(row[0]).date()
            end = datetime.fromisoformat(row[1]).date()
            periods_30.append((start, end))

        # Злиття послідовних періодів
        merged_100 = DataProcessor.merge_consecutive_periods(periods_100)
        merged_30 = DataProcessor.merge_consecutive_periods(periods_30)

        # Форматування
        formatted_100 = DataProcessor.format_periods_for_document(merged_100)
        formatted_30 = DataProcessor.format_periods_for_document(merged_30)

        # Оновлюємо таблицю periods
        cursor.execute("DELETE FROM periods WHERE servicemember_id = ?", (servicemember_id,))

        if formatted_100:
            cursor.execute("""
                INSERT INTO periods (servicemember_id, period_type, period_text)
                VALUES (?, '100', ?)
            """, (servicemember_id, formatted_100))

        if formatted_30:
            cursor.execute("""
                INSERT INTO periods (servicemember_id, period_type, period_text)
                VALUES (?, '30', ?)
            """, (servicemember_id, formatted_30))

        self.connection.commit()

    def get_unique_ranks(self) -> List[str]:
        """Отримати список унікальних звань"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT DISTINCT rank FROM servicemembers
            WHERE rank IS NOT NULL AND rank != ''
            ORDER BY rank
        """)
        return [row[0] for row in cursor.fetchall()]

    def get_available_months(self) -> List[str]:
        """Отримати список доступних місяців"""
        cursor = self.connection.cursor()
        cursor.execute("""
            SELECT DISTINCT month FROM service_records
            WHERE month IS NOT NULL
            ORDER BY month DESC
        """)
        return [row[0] for row in cursor.fetchall()]
