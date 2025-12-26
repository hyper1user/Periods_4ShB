"""
Стилі UI з мілітарі тематикою
"""

# Кольорова палітра (мілітарі тема)
COLORS = {
    "dark_green": "#2d4a2b",      # Темно-зелений
    "olive": "#556b2f",            # Оливковий
    "khaki": "#6b7c59",            # Хакі
    "light_green": "#8fbc8f",     # Світло-зелений
    "dark_gray": "#2f4f4f",       # Темно-сірий
    "light_gray": "#d3d3d3",      # Світло-сірий
    "white": "#ffffff",           # Білий
    "black": "#000000",           # Чорний
    "gold": "#ffd700",            # Золотий (для акцентів)
}

# Qt StyleSheet для мілітарі теми
MILITARY_STYLE = f"""
QMainWindow {{
    background-color: {COLORS['light_gray']};
}}

QLabel {{
    color: {COLORS['dark_green']};
    font-size: 14px;
    font-weight: bold;
}}

QLabel#title {{
    font-size: 24px;
    font-weight: bold;
    color: {COLORS['dark_green']};
    padding: 10px;
}}

QPushButton {{
    background-color: {COLORS['olive']};
    color: {COLORS['white']};
    border: 2px solid {COLORS['dark_green']};
    border-radius: 8px;
    padding: 12px 24px;
    font-size: 16px;
    font-weight: bold;
    min-width: 200px;
    min-height: 50px;
}}

QPushButton:hover {{
    background-color: {COLORS['khaki']};
    border: 2px solid {COLORS['gold']};
}}

QPushButton:pressed {{
    background-color: {COLORS['dark_green']};
    border: 2px solid {COLORS['light_green']};
}}

QPushButton:disabled {{
    background-color: {COLORS['dark_gray']};
    color: {COLORS['light_gray']};
    border: 2px solid {COLORS['dark_gray']};
}}

QComboBox {{
    background-color: {COLORS['white']};
    border: 2px solid {COLORS['olive']};
    border-radius: 5px;
    padding: 8px;
    font-size: 14px;
    min-height: 30px;
}}

QComboBox:hover {{
    border: 2px solid {COLORS['khaki']};
}}

QComboBox::drop-down {{
    border: none;
    background-color: {COLORS['olive']};
    width: 30px;
}}

QComboBox::down-arrow {{
    image: none;
    border: 2px solid {COLORS['white']};
    width: 8px;
    height: 8px;
    border-width: 0 2px 2px 0;
    transform: rotate(45deg);
}}

QCheckBox {{
    color: {COLORS['dark_green']};
    font-size: 14px;
    font-weight: bold;
    spacing: 8px;
}}

QCheckBox::indicator {{
    width: 20px;
    height: 20px;
    border: 2px solid {COLORS['olive']};
    border-radius: 4px;
    background-color: {COLORS['white']};
}}

QCheckBox::indicator:checked {{
    background-color: {COLORS['olive']};
    border: 2px solid {COLORS['dark_green']};
}}

QProgressBar {{
    border: 2px solid {COLORS['olive']};
    border-radius: 5px;
    text-align: center;
    background-color: {COLORS['white']};
    color: {COLORS['dark_green']};
    font-weight: bold;
}}

QProgressBar::chunk {{
    background-color: {COLORS['olive']};
    border-radius: 3px;
}}

QStatusBar {{
    background-color: {COLORS['khaki']};
    color: {COLORS['white']};
    font-weight: bold;
    font-size: 12px;
}}

QGroupBox {{
    border: 2px solid {COLORS['olive']};
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 10px;
    font-size: 14px;
    font-weight: bold;
    color: {COLORS['dark_green']};
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 5px;
    background-color: {COLORS['light_gray']};
}}

QMessageBox {{
    background-color: {COLORS['light_gray']};
}}

QMessageBox QPushButton {{
    min-width: 80px;
    min-height: 30px;
}}

QDialog {{
    background-color: {COLORS['light_gray']};
}}
"""


def get_military_style():
    """
    Повертає стиль для застосування до Qt додатку

    Returns:
        str: Qt StyleSheet
    """
    return MILITARY_STYLE
