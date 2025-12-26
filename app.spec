# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file для 1СБ 4ШБ Document Generator
"""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# Шлях до проекту
project_path = Path(SPECPATH)

# Збираємо всі файли openpyxl
openpyxl_datas, openpyxl_binaries, openpyxl_hiddenimports = collect_all('openpyxl')

a = Analysis(
    ['src/main.py'],
    pathex=[str(project_path)],
    binaries=openpyxl_binaries,
    datas=[
        # Шаблони документів
        ('templates', 'templates'),
        # Налаштування
        ('config', 'config'),
        # Додаткові дані (ЖБД, громади)
        ('Dodatky.xlsx', '.'),
        # Модулі програми
        ('src/gui', 'gui'),
        ('src/core', 'core'),
        ('src/utils', 'utils'),
    ] + openpyxl_datas,
    hiddenimports=[
        # PySide6 модулі
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        # Для роботи з Excel
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        # Для роботи з Word
        'docx',
        'docx.shared',
        'docx.enum.text',
        'docx.oxml',
        # Наші модулі
        'src.core.database',
        'src.core.excel_reader',
        'src.core.report_generator',
        'src.core.data_processor',
        'src.core.dodatky_reader',
        'src.core.migration',
        'src.core.updater',
        'src.gui.main_window',
        'src.gui.add_data_dialog',
        'src.gui.import_data_dialog',
        'src.gui.selection_dialog',
        'src.gui.styles',
        'src.gui.passport_data_dialog',
        'src.utils.date_utils',
        'src.utils.validators',
    ] + openpyxl_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Виключаємо непотрібні модулі для зменшення розміру
        'tkinter',
        'unittest',
        'test',
        'tests',
        'numpy',
        'pandas',
        'matplotlib',
        'scipy',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Periods_4SHB',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # GUI додаток без консолі
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Іконка програми
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Periods_4SHB',
)
