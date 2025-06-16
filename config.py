"""
Конфигурационные настройки для проверки адресов QazPost
"""
from pathlib import Path

# === ПУТИ К ФАЙЛАМ И ПАПКАМ ===
# Папки проекта
HTML_DIR = 'regions_html'
TABLES_DIR = 'tables'
BACKUP_DIR = 'backups'

# Файлы Excel
INPUT_EXCEL = 'tables/addresses.xlsx'
OUTPUT_EXCEL = 'tables/addresses_with_results.xlsx'
TEMPLATE_EXCEL = 'tables/template_addresses.xlsx'

# Логирование
LOG_FILE = 'address_check.log'

# Паттерны файлов
HTML_PATTERN = '*.html'

# === НАСТРОЙКИ EXCEL ===
# Строка с которой начинаются данные (0-based)
EXCEL_START_ROW = 8  # 9-я строка в Excel

# Колонки для чтения адресов
EXCEL_SETTLEMENT_COL = 'L'  # Населённый пункт
EXCEL_STREET_COL = 'M'      # Улица
EXCEL_HOUSE_COL = 'N'       # Дом

# Колонки для записи результатов
EXCEL_RESULT_COL = 'U'      # Статус проверки
EXCEL_DETAILS_COL = 'V'     # Найденный адрес

# === ПОРОГИ СОПОСТАВЛЕНИЯ ===
# Минимальное сходство для точного совпадения
SETTLEMENT_MATCH_THRESHOLD = 0.9  # Город/село
STREET_MATCH_THRESHOLD = 0.9      # Улица

# Минимальное сходство для частичного совпадения  
PARTIAL_MATCH_THRESHOLD = 0.4

# Веса для расчёта общего сходства
STREET_WEIGHT = 0.7  # Вес улицы
HOUSE_WEIGHT = 0.3   # Вес дома

# === НАСТРОЙКИ ЛОГИРОВАНИЯ ===
LOG_LEVEL_FILE = 'INFO'
LOG_LEVEL_CONSOLE = 'WARNING'
LOG_FORMAT = '%(asctime)s %(levelname)s: %(message)s'

# === CSS СЕЛЕКТОРЫ ДЛЯ ПАРСИНГА ===
# Селекторы для извлечения данных из HTML QazPost
OFFICE_CONTAINER_CLASS = 'DdeCNNHT'    # Контейнер офиса
ADDRESS_BLOCK_CLASS = '_3w4rWaD9'      # Блок с адресом

# === АВТОМАТИЧЕСКОЕ СОЗДАНИЕ ПАПОК ===
def ensure_directories():
    """Создаёт необходимые папки если их нет"""
    dirs = [HTML_DIR, TABLES_DIR, BACKUP_DIR]
    for dir_path in dirs:
        Path(dir_path).mkdir(exist_ok=True)

# Создаём папки при импорте модуля
ensure_directories()