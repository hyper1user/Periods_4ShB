"""
Модуль для перевірки та завантаження оновлень з GitHub
"""
import os
import sys
import json
import urllib.request
import urllib.error
import subprocess
import tempfile
from typing import Optional, Tuple
from dataclasses import dataclass

# Поточна версія програми
CURRENT_VERSION = "2.1.0"

# GitHub репозиторій (формат: username/repo)
GITHUB_REPO = "hyper1user/Periods_4ShB"


@dataclass
class ReleaseInfo:
    """Інформація про реліз"""
    version: str
    download_url: str
    release_url: str
    description: str
    published_at: str


def get_current_version() -> str:
    """Повертає поточну версію програми"""
    return CURRENT_VERSION


def parse_version(version: str) -> Tuple[int, ...]:
    """Парсить версію в кортеж чисел для порівняння"""
    # Видаляємо 'v' на початку якщо є
    version = version.lstrip('v')
    try:
        return tuple(int(x) for x in version.split('.'))
    except ValueError:
        return (0, 0, 0)


def is_newer_version(remote: str, local: str) -> bool:
    """Перевіряє чи remote версія новіша за local"""
    return parse_version(remote) > parse_version(local)


def check_for_updates(repo: str = None) -> Optional[ReleaseInfo]:
    """
    Перевіряє наявність оновлень на GitHub

    Args:
        repo: GitHub репозиторій (username/repo), якщо None - використовує GITHUB_REPO

    Returns:
        ReleaseInfo якщо є нова версія, None якщо немає або помилка
    """
    if repo is None:
        repo = GITHUB_REPO

    api_url = f"https://api.github.com/repos/{repo}/releases/latest"

    try:
        request = urllib.request.Request(
            api_url,
            headers={
                'User-Agent': 'Periods-4SHB-Updater',
                'Accept': 'application/vnd.github.v3+json'
            }
        )

        with urllib.request.urlopen(request, timeout=10) as response:
            data = json.loads(response.read().decode('utf-8'))

        version = data.get('tag_name', '').lstrip('v')
        if not version:
            return None

        # Перевіряємо чи версія новіша
        if not is_newer_version(version, CURRENT_VERSION):
            return None

        # Шукаємо ZIP файл в assets
        download_url = None
        for asset in data.get('assets', []):
            if asset['name'].endswith('.zip'):
                download_url = asset['browser_download_url']
                break

        # Якщо немає ZIP в assets - використовуємо zipball
        if not download_url:
            download_url = data.get('zipball_url')

        return ReleaseInfo(
            version=version,
            download_url=download_url,
            release_url=data.get('html_url', ''),
            description=data.get('body', ''),
            published_at=data.get('published_at', '')
        )

    except urllib.error.HTTPError as e:
        if e.code == 404:
            print(f"Репозиторій {repo} не знайдено або немає релізів")
        else:
            print(f"HTTP помилка: {e.code}")
        return None
    except urllib.error.URLError as e:
        print(f"Помилка з'єднання: {e.reason}")
        return None
    except Exception as e:
        print(f"Помилка при перевірці оновлень: {e}")
        return None


def download_update(release: ReleaseInfo, progress_callback=None) -> Optional[str]:
    """
    Завантажує оновлення

    Args:
        release: Інформація про реліз
        progress_callback: Функція для оновлення прогресу (current, total)

    Returns:
        Шлях до завантаженого файлу або None
    """
    if not release.download_url:
        return None

    try:
        # Створюємо тимчасовий файл
        temp_dir = tempfile.gettempdir()
        zip_path = os.path.join(temp_dir, f"periods_update_{release.version}.zip")

        request = urllib.request.Request(
            release.download_url,
            headers={'User-Agent': 'Periods-4SHB-Updater'}
        )

        with urllib.request.urlopen(request, timeout=60) as response:
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            chunk_size = 8192

            with open(zip_path, 'wb') as f:
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)

                    if progress_callback and total_size > 0:
                        progress_callback(downloaded, total_size)

        return zip_path

    except Exception as e:
        print(f"Помилка завантаження: {e}")
        return None


def get_app_directory() -> str:
    """Повертає директорію програми"""
    if getattr(sys, 'frozen', False):
        # Запущено як exe
        return os.path.dirname(sys.executable)
    else:
        # Запущено як скрипт
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_updater_path() -> str:
    """Повертає шлях до updater.exe"""
    app_dir = get_app_directory()

    # Спочатку шукаємо в _internal
    internal_path = os.path.join(app_dir, '_internal', 'updater.exe')
    if os.path.exists(internal_path):
        return internal_path

    # Потім в корені
    root_path = os.path.join(app_dir, 'updater.exe')
    if os.path.exists(root_path):
        return root_path

    return root_path  # Повертаємо очікуваний шлях


def run_updater(zip_path: str) -> bool:
    """
    Запускає updater.exe і закриває програму

    Args:
        zip_path: Шлях до ZIP з оновленням

    Returns:
        True якщо успішно запущено
    """
    updater_path = get_updater_path()
    app_dir = get_app_directory()

    if not os.path.exists(updater_path):
        print(f"Updater не знайдено: {updater_path}")
        return False

    try:
        # Запускаємо updater
        exe_name = "Periods_4SHB.exe"
        subprocess.Popen(
            [updater_path, zip_path, app_dir, exe_name],
            cwd=app_dir,
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )
        return True
    except Exception as e:
        print(f"Помилка запуску updater: {e}")
        return False


def open_release_page(release: ReleaseInfo):
    """Відкриває сторінку релізу в браузері"""
    import webbrowser
    if release.release_url:
        webbrowser.open(release.release_url)
