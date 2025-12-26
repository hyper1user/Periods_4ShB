"""
Обробка даних та ключова логіка злиття періодів
"""
from datetime import date
from typing import List, Tuple, Dict, Optional
from utils.date_utils import parse_period_string, format_period, is_consecutive


class DataProcessor:
    """
    Клас для обробки даних військовослужбовців та злиття періодів
    """

    @staticmethod
    def parse_periods(raw_periods: str) -> List[Tuple[date, date]]:
        """
        Парсить рядок з періодами у список кортежів дат

        Args:
            raw_periods: Рядок з періодами, розділеними переносами рядка
                        Формат: "з DD.MM.YYYY по DD.MM.YYYY"

        Returns:
            Список кортежів (start_date, end_date)
        """
        if not raw_periods:
            return []

        periods = []

        # Розділити по переносах рядків
        lines = raw_periods.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Парсити період з рядка
            period = parse_period_string(line)
            if period:
                periods.append(period)

        return periods

    @staticmethod
    def merge_consecutive_periods(periods: List[Tuple[date, date]]) -> List[Tuple[date, date]]:
        """
        Зливає послідовні періоди в один

        Алгоритм:
        1. Сортує періоди за датою початку
        2. Ітерує через відсортований список
        3. Якщо end_date[i] + 1 день == start_date[i+1]: об'єднує періоди
        4. Інакше: зберігає як окремі періоди

        Args:
            periods: Список кортежів (start_date, end_date)

        Returns:
            Список злитих періодів
        """
        if not periods:
            return []

        # Сортуємо періоди за датою початку
        sorted_periods = sorted(periods, key=lambda x: x[0])

        merged = []
        current_start, current_end = sorted_periods[0]

        for i in range(1, len(sorted_periods)):
            next_start, next_end = sorted_periods[i]

            # Перевіряємо чи періоди послідовні
            if is_consecutive(current_end, next_start):
                # Зливаємо періоди - розширюємо кінцеву дату
                current_end = next_end
            else:
                # Зберігаємо поточний період і починаємо новий
                merged.append((current_start, current_end))
                current_start, current_end = next_start, next_end

        # Додаємо останній період
        merged.append((current_start, current_end))

        return merged

    @staticmethod
    def format_periods_for_document(periods: List[Tuple[date, date]]) -> str:
        """
        Форматує періоди для вставки в документ

        Args:
            periods: Список кортежів (start_date, end_date)

        Returns:
            Рядок формату "з DD.MM.YYYY по DD.MM.YYYY, з DD.MM.YYYY по DD.MM.YYYY"
        """
        if not periods:
            return ""

        formatted_periods = []
        for start_date, end_date in periods:
            formatted_periods.append(format_period(start_date, end_date))

        return ", ".join(formatted_periods)

    @staticmethod
    def aggregate_servicemember_data(
        data_source,
        name: str,
        sheet_names: List[str] = None
    ) -> Optional[Dict]:
        """
        Агрегує всі дані для військовослужбовця з вказаних аркушів

        АДАПТОВАНИЙ метод - працює з обома джерелами даних (Excel або БД)

        Args:
            data_source: Екземпляр ExcelReader або DatabaseManager
            name: ПІБ військовослужбовця
            sheet_names: Список назв аркушів для обробки (тільки для Excel)

        Returns:
            Словник з агрегованими даними
        """
        # Визначити тип джерела даних
        from core.excel_reader import ExcelReader
        from core.database import DatabaseManager

        if isinstance(data_source, DatabaseManager):
            # Читання з БД - БД сам розраховує періоди
            return data_source.get_complete_data(name)
        elif isinstance(data_source, ExcelReader):
            # Існуюча логіка для Excel
            return DataProcessor._aggregate_from_excel(data_source, name, sheet_names)
        else:
            raise TypeError(f"Unknown data source type: {type(data_source)}")

    @staticmethod
    def _aggregate_from_excel(
        excel_reader,
        name: str,
        sheet_names: List[str]
    ) -> Optional[Dict]:
        """
        Агрегує дані з Excel (існуюча логіка)

        Args:
            excel_reader: Екземпляр ExcelReader
            name: ПІБ військовослужбовця
            sheet_names: Список назв аркушів для обробки

        Returns:
            Словник з агрегованими даними
        """
        # Отримати основну інформацію з аркуша Data
        info = excel_reader.get_servicemember_info_from_data(name)

        if not info:
            return None

        # Збираємо періоди окремо для 100% та 30%
        periods_100 = []
        periods_30 = []

        for sheet_name in sheet_names:
            data = excel_reader.get_servicemember_data(name, sheet_name)

            for row in data:
                if row.get("periods"):
                    periods = DataProcessor.parse_periods(row["periods"])

                    if sheet_name == "Періоди на 100":
                        periods_100.extend(periods)
                    elif sheet_name == "Періоди на 30":
                        periods_30.extend(periods)

        # Зливаємо послідовні періоди для кожного типу
        merged_100 = DataProcessor.merge_consecutive_periods(periods_100)
        merged_30 = DataProcessor.merge_consecutive_periods(periods_30)

        # Зливаємо всі періоди разом
        all_periods = periods_100 + periods_30
        merged_all = DataProcessor.merge_consecutive_periods(all_periods)

        # Форматуємо періоди для документу
        formatted_100 = DataProcessor.format_periods_for_document(merged_100)
        formatted_30 = DataProcessor.format_periods_for_document(merged_30)
        formatted_all = DataProcessor.format_periods_for_document(merged_all)

        return {
            "name": info["name"],
            "rank": info["rank"],
            "position": info["position"],
            "rnokpp": info["rnokpp"],
            "unit": info["unit"],
            "birth_date": info.get("birth_date"),
            "periods": formatted_all,  # Для Pilgova: {{ПЕРІОДИ}} містить всі періоди (100% + 30%)
            "periods_100": formatted_100,  # Для Only100: {{ПЕРІОДИ_100}} містить тільки періоди на 100%
            "periods_30": formatted_30,
            "periods_all": formatted_all,
            "periods_list": merged_all
        }

    @staticmethod
    def process_servicemembers_batch(
        data_source,
        names: List[str],
        sheet_names: List[str] = None
    ) -> List[Dict]:
        """
        Обробляє пакет військовослужбовців

        АДАПТОВАНИЙ метод - працює з обома джерелами даних (Excel або БД)

        Args:
            data_source: Екземпляр ExcelReader або DatabaseManager
            names: Список ПІБ військовослужбовців
            sheet_names: Список назв аркушів для обробки (тільки для Excel)

        Returns:
            Список словників з даними
        """
        results = []

        for name in names:
            data = DataProcessor.aggregate_servicemember_data(
                data_source,
                name,
                sheet_names
            )

            if data:
                results.append(data)

        return results
