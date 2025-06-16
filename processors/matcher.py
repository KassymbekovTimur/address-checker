"""
Система сопоставления адресов с базой данных QazPost
"""
import re
from difflib import SequenceMatcher
from utils.logger import setup_logger
import config

logger = setup_logger()

class AddressMatcher:
    """Класс для сопоставления адресов"""
    
    def __init__(self, offices_data):
        """
        Инициализация matcher'а
        
        Args:
            offices_data (dict): База данных офисов {settlement: [office_data, ...]}
        """
        self.offices_data = offices_data
        self.settlement_cache = {}  # Кэш для оптимизации поиска
        
        # Подготавливаем данные для быстрого поиска
        self._prepare_search_cache()
    
    def _prepare_search_cache(self):
        """Подготавливает кэш для быстрого поиска"""
        logger.info("Подготовка кэша для поиска...")
        
        for settlement, offices in self.offices_data.items():
            # Нормализуем название поселения для поиска
            normalized_settlement = self._normalize_text(settlement)
            
            if normalized_settlement not in self.settlement_cache:
                self.settlement_cache[normalized_settlement] = []
            
            # Добавляем все варианты этого поселения
            self.settlement_cache[normalized_settlement].extend(offices)
        
        logger.info(f"Кэш подготовлен: {len(self.settlement_cache)} поселений")
    
    def match_address(self, address_data):
        """
        Сопоставляет адрес с базой данных QazPost
        
        Args:
            address_data (dict): Данные адреса из Excel
            
        Returns:
            dict: Результат сопоставления
        """
        row_num = address_data['row_num']
        settlement = address_data['settlement']
        street = address_data['street']
        house = address_data['house']
        
        logger.debug(f"Строка {row_num}: Проверяем '{settlement}, {street}, {house}'")
        
        try:
            # 1. Ищем подходящие поселения
            matching_settlements = self._find_matching_settlements(settlement)
            
            if not matching_settlements:
                logger.debug(f"Строка {row_num}: Поселение '{settlement}' не найдено")
                return {
                    'row_num': row_num,
                    'status': 'Нет',
                    'details': f"Поселение '{settlement}' не найдено в базе QazPost"
                }
            
            # 2. Для каждого подходящего поселения ищем улицы
            best_match = None
            best_score = 0
            
            for settlement_match in matching_settlements:
                offices = settlement_match['offices']
                
                for office in offices:
                    # Сопоставляем улицу и дом
                    match_result = self._match_street_and_house(
                        street, house, office, settlement_match['score']
                    )
                    
                    if match_result['score'] > best_score:
                        best_score = match_result['score']
                        best_match = {
                            'office': office,
                            'settlement_score': settlement_match['score'],
                            'total_score': match_result['score'],
                            'details': match_result['details']
                        }
            
            # 3. Определяем статус на основе лучшего совпадения
            if best_match:
                result = self._determine_status(best_match, address_data)
                logger.debug(f"Строка {row_num}: {result['status']} (счёт: {best_score:.2f})")
                return result
            else:
                return {
                    'row_num': row_num,
                    'status': 'Нет',
                    'details': f"Улица '{street}' не найдена в поселении '{settlement}'"
                }
                
        except Exception as e:
            logger.error(f"Ошибка при сопоставлении строки {row_num}: {e}")
            return {
                'row_num': row_num,
                'status': 'Проверить',
                'details': f"Ошибка при проверке: {e}"
            }
    
    def _find_matching_settlements(self, settlement_name):
        """
        Находит подходящие поселения в базе данных
        
        Args:
            settlement_name (str): Название поселения для поиска
            
        Returns:
            list: Список подходящих поселений с оценками
        """
        normalized_input = self._normalize_text(settlement_name)
        matches = []
        
        for cached_settlement, offices in self.settlement_cache.items():
            similarity = self._calculate_similarity(normalized_input, cached_settlement)
            
            if similarity >= config.SETTLEMENT_MATCH_THRESHOLD:
                matches.append({
                    'name': cached_settlement,
                    'offices': offices,
                    'score': similarity
                })
        
        # Сортируем по убыванию сходства
        matches.sort(key=lambda x: x['score'], reverse=True)
        
        return matches
    
    def _match_street_and_house(self, street, house, office, settlement_score):
        """
        Сопоставляет улицу и дом с офисом
        
        Args:
            street (str): Улица из Excel
            house (str): Дом из Excel
            office (dict): Данные офиса из базы
            settlement_score (float): Оценка сходства поселения
            
        Returns:
            dict: Результат сопоставления с оценкой
        """
        office_street = office.get('street', '')
        office_house = office.get('house', '')
        
        # Нормализуем данные
        normalized_street = self._normalize_text(street)
        normalized_office_street = self._normalize_text(office_street)
        
        # Рассчитываем сходство улицы
        street_similarity = self._calculate_similarity(normalized_street, normalized_office_street)
        
        # Рассчитываем сходство дома
        house_similarity = self._calculate_house_similarity(house, office_house)
        
        # Общая оценка с весами
        total_score = (
            settlement_score * 0.3 +  # Вес поселения
            street_similarity * config.STREET_WEIGHT +
            house_similarity * config.HOUSE_WEIGHT
        )
        
        details = (f"Найден: {office['settlement']}, {office_street}, {office_house} "
                  f"(улица: {street_similarity:.2f}, дом: {house_similarity:.2f})")
        
        return {
            'score': total_score,
            'street_similarity': street_similarity,
            'house_similarity': house_similarity,
            'details': details
        }
    
    def _calculate_house_similarity(self, house1, house2):
        """
        Рассчитывает сходство номеров домов
        
        Args:
            house1 (str): Первый номер дома
            house2 (str): Второй номер дома
            
        Returns:
            float: Оценка сходства (0.0 - 1.0)
        """
        if not house1 or not house2:
            return 0.0
        
        # Извлекаем числовые части
        num1 = self._extract_house_number(house1)
        num2 = self._extract_house_number(house2)
        
        # Точное совпадение
        if house1.strip().lower() == house2.strip().lower():
            return 1.0
        
        # Совпадение числовых частей
        if num1 and num2 and num1 == num2:
            return 0.9
        
        # Частичное сходство строк
        return self._calculate_similarity(house1, house2)
    
    def _extract_house_number(self, house_str):
        """
        Извлекает основной номер дома из строки
        
        Args:
            house_str (str): Строка с номером дома
            
        Returns:
            str: Основной номер дома или None
        """
        if not house_str:
            return None
        
        # Ищем первое число в строке
        match = re.search(r'\d+', str(house_str))
        return match.group() if match else None
    
    def _determine_status(self, best_match, address_data):
        """
        Определяет статус проверки на основе лучшего совпадения
        
        Args:
            best_match (dict): Лучшее совпадение
            address_data (dict): Исходные данные адреса
            
        Returns:
            dict: Результат с статусом
        """
        total_score = best_match['total_score']
        office = best_match['office']
        
        # Определяем статус
        if total_score >= 0.9:
            status = 'Да'
        elif total_score >= config.PARTIAL_MATCH_THRESHOLD:
            status = 'Проверить'
        else:
            status = 'Нет'
        
        # Формируем детальную информацию
        details = best_match['details']
        
        return {
            'row_num': address_data['row_num'],
            'status': status,
            'details': details
        }
    
    def _normalize_text(self, text):
        """
        Нормализует текст для сравнения
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Нормализованный текст
        """
        if not text:
            return ''
        
        # Приводим к нижнему регистру
        text = str(text).lower().strip()
        
        # Убираем лишние пробелы
        text = ' '.join(text.split())
        
        # Убираем типичные префиксы и суффиксы
        text = re.sub(r'^(г\.|город|с\.|село|пос\.|посёлок|ул\.|улица|пр\.|проспект|мкр\.|микрорайон)\s*', '', text)
        text = re.sub(r'\s*(г\.|город|с\.|село|пос\.|посёлок|ул\.|улица|пр\.|проспект|мкр\.|микрорайон)$', '', text)
        
        # Заменяем синонимы
        replacements = {
            'проспект': 'пр',
            'улица': 'ул',
            'микрорайон': 'мкр',
            'переулок': 'пер',
            'бульвар': 'бул'
        }
        
        for old, new in replacements.items():
            text = text.replace(old, new)
        
        return text.strip()
    
    def _calculate_similarity(self, text1, text2):
        """
        Рассчитывает сходство между двумя текстами
        
        Args:
            text1 (str): Первый текст
            text2 (str): Второй текст
            
        Returns:
            float: Оценка сходства (0.0 - 1.0)
        """
        if not text1 or not text2:
            return 0.0
        
        if text1 == text2:
            return 1.0
        
        # Используем SequenceMatcher для расчёта сходства
        return SequenceMatcher(None, text1, text2).ratio()