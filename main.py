"""
Система проверки адресов QazPost
Главный модуль приложения
"""
from utils.logger import setup_logger
from parsers.html_parser import HTMLParser
from processors.excel_processor import ExcelProcessor
from processors.matcher import AddressMatcher
import config

logger = setup_logger()

def main():
    """Основная функция приложения"""
    try:
        
        # 1. Парсинг HTML файлов
        logger.info("Этап 1: Парсинг HTML файлов...")
        html_parser = HTMLParser()
        offices_data = html_parser.parse_html_files()
        
        if not offices_data:
            raise Exception("Не найдено данных в HTML файлах. Проверьте папку regions_html/")
        
        # Статистика базы данных
        stats = html_parser.get_statistics()
        logger.info(f"База данных загружена: {stats['total_settlements']} поселений, "
                   f"{stats['total_offices']} офисов")
        
        # 2. Обработка Excel файла
        logger.info("Этап 2: Загрузка Excel файла...")
        excel_processor = ExcelProcessor()
        workbook, worksheet = excel_processor.load_workbook()
        
        total_rows = excel_processor.get_total_rows()
        logger.info(f"К обработке: {total_rows} записей")
        
        # 3. Сопоставление адресов
        logger.info("Этап 3: Сопоставление адресов...")
        matcher = AddressMatcher(offices_data)
        
        results = []
        processed_count = 0
        
        # Обрабатываем строки начиная с EXCEL_START_ROW + 1 (1-based indexing)
        for row_num in range(config.EXCEL_START_ROW + 1, worksheet.max_row + 1):
            address_data = excel_processor.read_address_row(row_num)
            
            if address_data:
                result = matcher.match_address(address_data)
                results.append(result)
                processed_count += 1
                
                # Логируем прогресс каждые 100 строк
                if processed_count % 100 == 0:
                    logger.info(f"Обработано: {processed_count}/{total_rows}")
        
        logger.info(f"Сопоставление завершено. Обработано записей: {processed_count}")
        
        # 4. Сохранение результатов
        logger.info("Этап 4: Сохранение результатов...")
        excel_processor.save_results(results)
        
        # 5. Статистика
        logger.info("Этап 5: Подготовка статистики...")
        print_statistics(results)
        
        # Закрываем Excel
        excel_processor.close()
        
        logger.info("=== ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО ===")
        print(f"\n✅ Обработка завершена!")
        print(f"📄 Результаты сохранены: {config.OUTPUT_EXCEL}")
        print(f"📋 Подробные логи: {config.LOG_FILE}")
        
    except FileNotFoundError as e:
        logger.error(f"Файл не найден: {e}")
        print(f"❌ Ошибка: {e}")
        print("Убедитесь, что файл addresses.xlsx находится в папке tables/")
        
    except PermissionError as e:
        logger.error(f"Ошибка доступа: {e}")
        print(f"❌ Ошибка: {e}")
        print("Закройте Excel файл и повторите попытку")
        
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
        print(f"❌ Произошла ошибка: {e}")
        print(f"Проверьте лог-файл: {config.LOG_FILE}")
        raise

def print_statistics(results):
    """
    Выводит статистику результатов
    
    Args:
        results (list): Список результатов проверки
    """
    if not results:
        print("Нет данных для статистики")
        return
    
    # Подсчитываем статистику
    stats = {'Да': 0, 'Проверить': 0, 'Нет': 0}
    
    for result in results:
        if result and 'status' in result:
            status = result['status']
            if status in stats:
                stats[status] += 1
    
    total = sum(stats.values())
    
    # Выводим результаты
    print("\n" + "="*50)
    print("📊 СТАТИСТИКА РЕЗУЛЬТАТОВ")
    print("="*50)
    
    for status, count in stats.items():
        percentage = (count / total * 100) if total > 0 else 0
        emoji = {'Да': '✅', 'Проверить': '⚠️', 'Нет': '❌'}[status]
        
        message = f"{emoji} {status}: {count} ({percentage:.1f}%)"
        print(message)
        logger.info(message)
    
    print("-" * 50)
    summary = f"📋 Всего обработано: {total} записей"
    print(summary)
    logger.info(summary)

if __name__ == '__main__':
    main()