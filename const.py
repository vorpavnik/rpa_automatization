import os

# Базовые пути
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Пути к файлам
IMAGE_PATH = os.path.join(SCRIPT_DIR, "ssg.bmp")
LOG_FILE = os.path.join(SCRIPT_DIR, "app.log")

# Настройки приложения
APP_TITLE = "Автоматизация АСУ ПРИГ"
APP_SIZE = "900x700"
EXCEL_FILE = "График обслуживания поездов МВПС.xlsx"
