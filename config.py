# config.py
import os
from pathlib import Path
from datetime import datetime, timedelta

# Базовые пути
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "эксельки"
OUTPUT_DIR = BASE_DIR / "вывод"
AD_EXPORT_DIR = INPUT_DIR / "AD"
SHTAT_DIR = INPUT_DIR / "штатка"
KONTUR_DIR = INPUT_DIR / "эдо_контур"
DIADOC_DIR = INPUT_DIR / "эдо_диадок"
ONEC_DIR = INPUT_DIR / "1С"

# Создаем директории, если они не существуют
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
AD_EXPORT_DIR.mkdir(exist_ok=True)
SHTAT_DIR.mkdir(exist_ok=True)
KONTUR_DIR.mkdir(exist_ok=True)
DIADOC_DIR.mkdir(exist_ok=True)
ONEC_DIR.mkdir(exist_ok=True)

# Настройка актуальности файлов (в днях)
MAX_FILE_AGE_DAYS = 30

# Генерация имени файла с датой и временеи
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = OUTPUT_DIR / f"результат_обработки_{current_time}.xlsx"

# Файлы сотрудников и ГПХ
EMPLOYEES_FILE = AD_EXPORT_DIR / "сотрудники.txt"
GPH_FILE = AD_EXPORT_DIR / "ГПХ.txt"

# Файлы ЭДО
KONTUR_FILE = KONTUR_DIR / "Контур.xlsx"
DIADOC_FILE = DIADOC_DIR / "Выгрузка_SBINV-39662.xlsx"

# Настройки обработки Excel
SHEET_NAME = "сравнение пользователей"
COMPARISON_SHEET = "сравнение AD и Штатки"
KONTUR_SHEET = "Контур данные"
DIADOC_SHEET = "Диадок данные"
ONEC_SHEET = "1С данные"
MAX_ROWS = 10000  # Увеличили лимит строк
RED_COLOR = (255, 199, 206)  # RGB для красного цвета
YELLOW_COLOR = (255, 235, 156)  # RGB для желтого цвета