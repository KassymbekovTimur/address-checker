"""
Парсер HTML файлов для извлечения адресов офисов QazPost
"""
import re
import glob
from pathlib import Path
from bs4 import BeautifulSoup
from utils.logger import setup_logger
import config

logger = setup_logger()

class HTMLParser:
    """Парсер HTML файлов QazPost"""
    
    def __init__(self):
        """Инициализация парсера"""
        self.offices_data = {}  # {settlement: [office_data, ...]}
    
    def parse_html_files(self):
        """
        Парсит все HTML файлы в папке regions_html
        
        Returns:
            dict: Словарь с данными офисов {settlement: [offices]}
        """
        html_files = glob.glob(f"{config.HTML_DIR}/{config.HTML_PATTERN}")
        
        if not html_files:
            logger.warning(f"HTML файлы не найдены в папке {config.HTML_DIR}")
            return {}
        
        logger.info(f"Найдено HTML файлов: {len(html_files)}")
        
        total_offices = 0
        for html_file in html_files:
            try:
                offices_count = self._parse_single_file(html_file)
                total_offices += offices_count
                logger.info(f"Файл {Path(html_file).name}: извлечено {offices_count} офисов")
                
            except Exception as e:
                logger.error(f"Ошибка при парсинге {html_file}: {e}")
        
        logger.info(f"Всего извлечено офисов: {total_offices}")
        logger.info(f"Поселений в базе: {len(self.offices_data)}")
        
        return self.offices_data
    
    def _parse_single_file(self, html_file):
        """
        Парсит один HTML файл
        
        Args:
            html_file (str): Путь к HTML файлу
            
        Returns:
            int: Количество найденных офисов
        """
        with open(html_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        soup = BeautifulSoup(content, 'html.parser')
        
        # Ищем контейнеры офисов
        office_containers = soup.find_all('div', class_=config.OFFICE_CONTAINER_CLASS)
        
        offices_count = 0
        for container in office_containers:
            try:
                office_data = self._extract_office_data(container)
                if office_data:
                    self._add_office_to_database(office_data)
                    offices_count += 1
                    
            except Exception as e:
                logger.debug(f"Ошибка при извлечении данных офиса: {e}")
        
        return offices_count
    
    def _extract_office_data(self, container):
        """
        Извлекает данные офиса из HTML контейнера
        
        Args:
            container: BeautifulSoup элемент контейнера офиса
            
        Returns:
            dict: Данные офиса или None если не удалось извлечь
        """
        # Ищем блок с адресом
        address_block = container.find('div', class_=config.ADDRESS_BLOCK_CLASS)
        if not address_block:
            return None
        
        address_text = address_block.get_text(strip=True)
        if not address_text:
            return None
        
        # Парсим адрес
        parsed_address = self._parse_address_string(address_text)
        if not parsed_address:
            return None
        
        return {
            'full_address': address_text,
            'settlement': parsed_address['settlement'],
            'street': parsed_address['street'],
            'house': parsed_address['house']
        }
    
    def _parse_address_string(self, address_text):
        """
        Парсит строку адреса и извлекает компоненты
        
        Args:
            address_text (str): Строка с адресом
            
        Returns:
            dict: Компоненты адреса или None
        """
        # Убираем лишние пробелы и приводим к нижнему регистру для анализа
        clean_address = ' '.join(address_text.split())
        
        # Паттерн для парсинга адреса
        # Примеры: "г. Алматы, ул. Абая, д. 150", "Астана, пр. Кунаева, 12"
        patterns = [
            # г. Город, ул. Улица, д. Дом
            r'(?:г\.\s*)?([^,]+),\s*(?:ул\.|пр\.|мкр\.)\s*([^,]+),\s*(?:д\.\s*)?(.+)',
            # Город, ул. Улица, Дом  
            r'([^,]+),\s*(?:ул\.|пр\.|мкр\.)\s*([^,]+),\s*(.+)',
            # г. Город ул. Улица д. Дом (без запятых)
            r'(?:г\.\s*)?([^,]+)\s+(?:ул\.|пр\.|мкр\.)\s*([^,]+)\s+(?:д\.\s*)?(.+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, clean_address, re.IGNORECASE)
            if match:
                settlement = match.group(1).strip()
                street = match.group(2).strip()  
                house = match.group(3).strip()
                
                # Очищаем данные
                settlement = self._clean_settlement_name(settlement)
                street = self._clean_street_name(street)
                house = self._clean_house_number(house)
                
                if settlement and street and house:
                    return {
                        'settlement': settlement,
                        'street': street,
                        'house': house
                    }
        
        logger.debug(f"Не удалось распарсить адрес: {address_text}")
        return None
    
    def _clean_settlement_name(self, settlement):
        """Очищает название населённого пункта"""
        settlement = settlement.strip()
        
        # Убираем префиксы типа "г.", "с.", "пос."
        settlement = re.sub(r'^(г\.|город|с\.|село|пос\.|посёлок)\s*', '', settlement, flags=re.IGNORECASE)
        
        return settlement.strip() if settlement else None
    
    def _clean_street_name(self, street):
        """Очищает название улицы"""
        street = street.strip()
        
        # Добавляем стандартные сокращения если их нет
        if not re.match(r'^(ул\.|пр\.|мкр\.)', street, re.IGNORECASE):
            # Пытаемся определить тип улицы
            if any(word in street.lower() for word in ['проспект', 'avenue']):
                street = f"пр. {street}"
            elif any(word in street.lower() for word in ['микрорайон', 'мкр']):
                street = f"мкр. {street}"
            else:
                street = f"ул. {street}"
        
        return street if street else None
    
    def _clean_house_number(self, house):
        """Очищает номер дома"""
        house = house.strip()
        
        # Убираем префиксы "д.", "дом"  
        house = re.sub(r'^(д\.|дом)\s*', '', house, flags=re.IGNORECASE)
        
        return house.strip() if house else None
    
    def _add_office_to_database(self, office_data):
        """
        Добавляет офис в базу данных
        
        Args:
            office_data (dict): Данные офиса
        """
        settlement = office_data['settlement'].lower()
        
        if settlement not in self.offices_data:
            self.offices_data[settlement] = []
        
        self.offices_data[settlement].append(office_data)
    
    def get_statistics(self):
        """
        Возвращает статистику по базе данных
        
        Returns:
            dict: Статистика
        """
        total_offices = sum(len(offices) for offices in self.offices_data.values())
        
        return {
            'total_settlements': len(self.offices_data),
            'total_offices': total_offices,
            'settlements': list(self.offices_data.keys())
        }