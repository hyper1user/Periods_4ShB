"""
Допоміжні функції для роботи з датами
"""
from datetime import date, datetime, timedelta
from typing import Optional
import re


def parse_ukrainian_date(date_str: str) -> Optional[date]:
    """
    Парсить дату з українського формату DD.MM.YYYY

    Args:
        date_str: Рядок з датою у форматі DD.MM.YYYY

    Returns:
        date object або None якщо парсинг не вдався
    """
    if not date_str or not isinstance(date_str, str):
        return None

    # Видаляємо зайві пробіли
    date_str = date_str.strip()

    # Спроба розпарсити у форматі DD.MM.YYYY
    try:
        return datetime.strptime(date_str, "%d.%m.%Y").date()
    except ValueError:
        pass

    # Спроба розпарсити у форматі YYYY-MM-DD
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        pass

    return None


def format_date_ukrainian(date_obj: date) -> str:
    """
    Форматує дату в український формат DD.MM.YYYY

    Args:
        date_obj: Об'єкт дати

    Returns:
        Рядок у форматі DD.MM.YYYY
    """
    if not date_obj:
        return ""

    return date_obj.strftime("%d.%m.%Y")


def is_consecutive(end_date: date, start_date: date) -> bool:
    """
    Перевіряє чи є дві дати послідовними (різниця 1 день)

    Args:
        end_date: Кінцева дата першого періоду
        start_date: Початкова дата другого періоду

    Returns:
        True якщо дати послідовні, False інакше
    """
    if not end_date or not start_date:
        return False

    # Перевіряємо чи наступний день після end_date дорівнює start_date
    return (end_date + timedelta(days=1)) == start_date


def calculate_total_days(periods: list[tuple[date, date]]) -> int:
    """
    Підраховує загальну кількість днів у списку періодів

    Args:
        periods: Список кортежів (start_date, end_date)

    Returns:
        Загальна кількість днів
    """
    total_days = 0

    for start, end in periods:
        if start and end:
            # +1 бо включаємо обидва дні
            total_days += (end - start).days + 1

    return total_days


def parse_period_string(period_str: str) -> Optional[tuple[date, date]]:
    """
    Парсить рядок періоду формату "з DD.MM.YYYY по DD.MM.YYYY"

    Args:
        period_str: Рядок з періодом

    Returns:
        Кортеж (start_date, end_date) або None
    """
    if not period_str or not isinstance(period_str, str):
        return None

    # Регулярний вираз для пошуку дат у форматі DD.MM.YYYY
    date_pattern = r'(\d{2}\.\d{2}\.\d{4})'
    dates = re.findall(date_pattern, period_str)

    if len(dates) >= 2:
        start_date = parse_ukrainian_date(dates[0])
        end_date = parse_ukrainian_date(dates[1])

        if start_date and end_date:
            return (start_date, end_date)

    return None


def format_period(start_date: date, end_date: date) -> str:
    """
    Форматує період у вигляді "з DD.MM.YYYY по DD.MM.YYYY"

    Args:
        start_date: Початкова дата
        end_date: Кінцева дата

    Returns:
        Відформатований рядок періоду
    """
    if not start_date or not end_date:
        return ""

    start_str = format_date_ukrainian(start_date)
    end_str = format_date_ukrainian(end_date)

    return f"з {start_str} по {end_str}"
