"""
Updater - окрема програма для оновлення Periods_4SHB
Запускається основною програмою, чекає її закриття, замінює файли і запускає нову версію
"""
import sys
import os
import time
import shutil
import zipfile
import subprocess
import psutil


def wait_for_process_exit(process_name: str, timeout: int = 30):
    """Чекає поки процес закриється"""
    print(f"Очікування закриття {process_name}...")

    start_time = time.time()
    while time.time() - start_time < timeout:
        found = False
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                    found = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        if not found:
            print(f"{process_name} закрито")
            return True

        time.sleep(0.5)

    print(f"Таймаут очікування {process_name}")
    return False


def extract_update(zip_path: str, target_dir: str):
    """Розпаковує оновлення"""
    print(f"Розпакування {zip_path}...")

    # Створюємо тимчасову папку для розпакування
    temp_extract_dir = os.path.join(os.path.dirname(zip_path), "_update_temp")

    # Видаляємо стару тимчасову папку якщо є
    if os.path.exists(temp_extract_dir):
        shutil.rmtree(temp_extract_dir)

    # Розпаковуємо
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_extract_dir)

    # Знаходимо кореневу папку в архіві (може бути Periods_4SHB_Portable_v2.x.x)
    extracted_items = os.listdir(temp_extract_dir)
    if len(extracted_items) == 1 and os.path.isdir(os.path.join(temp_extract_dir, extracted_items[0])):
        source_dir = os.path.join(temp_extract_dir, extracted_items[0])
    else:
        source_dir = temp_extract_dir

    print(f"Копіювання файлів з {source_dir} в {target_dir}...")

    # Копіюємо файли
    for item in os.listdir(source_dir):
        source_path = os.path.join(source_dir, item)
        target_path = os.path.join(target_dir, item)

        # Пропускаємо data.db, output, config - зберігаємо дані користувача
        if item in ['data.db', 'output', 'config']:
            print(f"  Пропускаємо {item} (дані користувача)")
            continue

        # Пропускаємо updater.exe якщо він зараз виконується
        if item == 'updater.exe':
            print(f"  Пропускаємо {item} (оновлюється)")
            continue

        try:
            if os.path.isdir(source_path):
                if os.path.exists(target_path):
                    shutil.rmtree(target_path)
                shutil.copytree(source_path, target_path)
            else:
                shutil.copy2(source_path, target_path)
            print(f"  ✓ {item}")
        except Exception as e:
            print(f"  ✗ {item}: {e}")

    # Видаляємо тимчасову папку
    try:
        shutil.rmtree(temp_extract_dir)
    except:
        pass

    # Видаляємо ZIP
    try:
        os.remove(zip_path)
    except:
        pass

    print("Оновлення завершено!")


def main():
    """Головна функція updater"""
    print("=" * 50)
    print("Periods_4SHB Updater")
    print("=" * 50)

    # Отримуємо аргументи
    if len(sys.argv) < 3:
        print("Використання: updater.exe <zip_path> <target_dir> [exe_name]")
        print("Аргументи:")
        print("  zip_path   - шлях до ZIP з оновленням")
        print("  target_dir - папка програми")
        print("  exe_name   - назва exe для запуску (опціонально)")
        input("Натисніть Enter для виходу...")
        sys.exit(1)

    zip_path = sys.argv[1]
    target_dir = sys.argv[2]
    exe_name = sys.argv[3] if len(sys.argv) > 3 else "Periods_4SHB.exe"

    print(f"ZIP: {zip_path}")
    print(f"Папка: {target_dir}")
    print(f"Програма: {exe_name}")
    print()

    # Перевіряємо чи існує ZIP
    if not os.path.exists(zip_path):
        print(f"Помилка: файл {zip_path} не знайдено!")
        input("Натисніть Enter для виходу...")
        sys.exit(1)

    # Чекаємо закриття основної програми
    wait_for_process_exit(exe_name)

    # Невелика пауза для надійності
    time.sleep(1)

    # Розпаковуємо оновлення
    try:
        extract_update(zip_path, target_dir)
    except Exception as e:
        print(f"Помилка при оновленні: {e}")
        input("Натисніть Enter для виходу...")
        sys.exit(1)

    # Запускаємо оновлену програму
    exe_path = os.path.join(target_dir, exe_name)
    if os.path.exists(exe_path):
        print(f"Запуск {exe_name}...")
        subprocess.Popen([exe_path], cwd=target_dir)
    else:
        print(f"Увага: {exe_path} не знайдено")

    print()
    print("Updater завершив роботу")
    time.sleep(2)


if __name__ == "__main__":
    main()
