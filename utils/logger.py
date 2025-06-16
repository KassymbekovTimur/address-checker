import logging
import sys
import config

def setup_logger():
    """
    Настраивает и возвращает logger для приложения
    
    Returns:
        logging.Logger: Настроенный logger
    """
    # Создаём logger
    logger = logging.getLogger('address_checker')
    logger.setLevel(logging.INFO)
    
    # Очищаем существующие handlers
    logger.handlers.clear()
    
    # Форматтер для логов
    formatter = logging.Formatter(config.LOG_FORMAT)
    
    # === FILE HANDLER ===
    file_handler = logging.FileHandler(
        config.LOG_FILE, 
        mode='w',  # Перезаписываем лог при каждом запуске
        encoding='utf-8'
    )
    file_handler.setLevel(getattr(logging, config.LOG_LEVEL_FILE))
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # === CONSOLE HANDLER ===
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, config.LOG_LEVEL_CONSOLE))
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger