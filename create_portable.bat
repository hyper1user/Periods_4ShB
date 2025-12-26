@echo off
chcp 65001 >nul
:: Створення portable версії для тестування
:: Періоди 4ШБ

echo ================================================================================
echo СТВОРЕННЯ PORTABLE ВЕРСІЇ "Періоди 4ШБ"
echo ================================================================================
echo.

:: Назва архіву
set PORTABLE_NAME=Periods_4SHB_Portable_v2.1.0
set DIST_DIR=dist\Periods_4SHB
set PORTABLE_DIR=portable\%PORTABLE_NAME%

echo [1/4] Перевірка наявності зібраного exe...
if not exist "%DIST_DIR%\Periods_4SHB.exe" (
    echo [ПОМИЛКА] Exe файл не знайдено!
    echo Спочатку зберіть програму: pyinstaller app.spec --clean -y
    pause
    exit /b 1
)
echo ✓ Exe знайдено

echo.
echo [2/4] Створення структури portable...
if exist "portable" rmdir /s /q "portable"
mkdir "%PORTABLE_DIR%"
mkdir "%PORTABLE_DIR%\templates"
mkdir "%PORTABLE_DIR%\output"

echo.
echo [3/4] Копіювання файлів...

:: Копіюємо exe
copy /Y "%DIST_DIR%\Periods_4SHB.exe" "%PORTABLE_DIR%\" >nul
echo   ✓ Periods_4SHB.exe

:: Копіюємо _internal
xcopy "%DIST_DIR%\_internal" "%PORTABLE_DIR%\_internal" /E /I /Y /Q >nul
echo   ✓ _internal\ (всі залежності)

:: Копіюємо шаблони якщо є
if exist "templates" (
    xcopy "templates" "%PORTABLE_DIR%\templates" /E /I /Y /Q >nul
    echo   ✓ templates\
)

:: Створюємо config для portable версії (БЕЗ абсолютних шляхів)
mkdir "%PORTABLE_DIR%\config"
if exist "config\settings_portable.json" (
    copy /Y "config\settings_portable.json" "%PORTABLE_DIR%\config\settings.json" >nul
    echo   ✓ config\ (portable налаштування)
) else (
    echo   [УВАГА] settings_portable.json не знайдено
)

:: Копіюємо Dodatky.xlsx якщо є
if exist "Dodatky.xlsx" (
    copy /Y "Dodatky.xlsx" "%PORTABLE_DIR%\_internal\Dodatky.xlsx" >nul
    echo   ✓ Dodatky.xlsx
)

:: Копіюємо data.db якщо є (БД з готовими даними)
if exist "data.db" (
    copy /Y "data.db" "%PORTABLE_DIR%\data.db" >nul
    echo   ✓ data.db (база даних з даними)
) else (
    echo   [INFO] data.db не знайдено - буде створено при першому запуску
)

:: Копіюємо updater.exe для автооновлення
if exist "dist\updater.exe" (
    copy /Y "dist\updater.exe" "%PORTABLE_DIR%\_internal\updater.exe" >nul
    echo   ✓ updater.exe (автооновлення)
) else (
    echo   [УВАГА] updater.exe не знайдено - збудуйте: pyinstaller updater.spec
)

:: Створюємо README для portable версії
echo Створення README.txt...
(
echo ================================================================================
echo                        Періоди 4ШБ - Portable версія
echo ================================================================================
echo.
echo PORTABLE ВЕРСІЯ - НЕ ПОТРЕБУЄ ВСТАНОВЛЕННЯ
echo.
echo Ця версія програми може працювати з будь-якої папки без встановлення.
echo Просто розпакуйте архів і запустіть Periods_4SHB.exe
echo.
echo СТРУКТУРА:
echo   Periods_4SHB.exe       - Головна програма
echo   _internal\             - Системні файли та залежності
echo   _internal\updater.exe  - Програма автооновлення
echo   templates\             - Шаблони документів Word
echo   output\                - Згенеровані документи (створюється автоматично^)
echo   data.db                - База даних (створюється при першому запуску^)
echo.
echo ПЕРШИЙ ЗАПУСК:
echo   1. Запустіть Periods_4SHB.exe
echo   2. Якщо БД вже включена - одразу генеруйте документи!
echo   3. Для додавання нових в/с або періодів - використовуйте кнопки в програмі
echo   4. Excel імпорт - опціонально, для масового завантаження
echo.
echo ВАЖЛИВО:
echo   - Не видаляйте папку _internal\
echo   - База даних зберігається у цій же папці
echo   - Документи зберігаються в output\
echo   - Для оновлення натисніть кнопку "Оновлення" в програмі
echo.
echo ПЕРЕВАГИ PORTABLE ВЕРСІЇ:
echo   ✓ Не потребує встановлення
echo   ✓ Можна запускати з флешки
echo   ✓ Легко переносити між комп'ютерами
echo   ✓ Всі дані в одній папці
echo.
echo ВЕРСІЯ: 2.1.0 (з автооновленням)
echo ================================================================================
) > "%PORTABLE_DIR%\README.txt"
echo   ✓ README.txt

echo.
echo [4/4] Створення ZIP архіву...

:: Перевірка наявності PowerShell
powershell -Command "Get-Command Compress-Archive" >nul 2>&1
if %errorlevel% equ 0 (
    powershell -Command "Compress-Archive -Path '%PORTABLE_DIR%\*' -DestinationPath 'portable\%PORTABLE_NAME%.zip' -Force"
    if %errorlevel% equ 0 (
        echo ✓ ZIP архів створено
    ) else (
        echo [УВАГА] Не вдалося створити ZIP архів
    )
) else (
    echo [УВАГА] PowerShell недоступний, ZIP не створено
    echo Ви можете створити архів вручну з папки: %PORTABLE_DIR%
)

echo.
echo ================================================================================
echo ✓ PORTABLE ВЕРСІЯ ГОТОВА!
echo ================================================================================
echo.
echo Папка: portable\%PORTABLE_NAME%\
if exist "portable\%PORTABLE_NAME%.zip" (
    echo Архів: portable\%PORTABLE_NAME%.zip
)
echo.
echo ТЕСТУВАННЯ БЕЗ ВСТАНОВЛЕННЯ:
echo   1. Перейдіть в папку: portable\%PORTABLE_NAME%\
echo   2. Запустіть: Periods_4SHB.exe
echo   3. Програма працюватиме без встановлення!
echo.
echo РОЗПОВСЮДЖЕННЯ:
echo   - Відправте ZIP архів користувачам
echo   - Вони просто розпаковують і запускають
echo   - Не потрібні права адміністратора
echo.
pause
