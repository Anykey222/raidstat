from thefuzz import process, fuzz
import re
import logging

class Matcher:
    def __init__(self, known_names=None):
        self.known_names = known_names or []
        self.replacements = {}
        self.logger = logging.getLogger(__name__)
        # Regex for valid names (Cyrillic/Latin)
        self.name_pattern = re.compile(r"^\W?([A-ZА-ЯЁ][a-zа-яё]+)(?:\.{0,2}[^a-zа-я]?|[^a-zа-я]?\.{0,2}|[^a-zа-я]{0,2}\.?)?$", re.DOTALL)

    def set_known_names(self, names):
        self.known_names = names

    def load_replacements(self, file_path="Замены.txt"):
        """Загрузка замен из файла."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    parts = line.strip().split(maxsplit=1)
                    if len(parts) == 2:
                        self.replacements[parts[0]] = parts[1]
        except Exception as e:
            self.logger.warning(f"Не удалось загрузить замены: {e}")

    def apply_replacements(self, text):
        if not text:
            return text
        
        # Проверка регулярных выражений для замен
        for pattern, replacement in self.replacements.items():
            if re.match(pattern, text):
                return replacement
        return text

    def smart_match(self, raw_text):
        """
        Возвращает: (matched_name, score, type_code)
        Коды типов: 
        0 = 100% совпадение
        1 = Нечеткое совпадение
        2 = Новое имя (высокая уверенность регулярного выражения)
        3 = Ошибка/Неизвестно (всегда возвращается, если совпадение не найдено)
        4 = Применена замена
        """
        if not raw_text or len(raw_text) <= 2:
            return None, 0, 3

        # 1. Применяем замены
        text_replaced = self.apply_replacements(raw_text)
        if text_replaced != raw_text:
            return text_replaced, 100, 4
        
        # 2. Проверка регулярным выражением
        match = self.name_pattern.search(raw_text)
        if match:
            clean_name = match.group(1)
            
            if not self.known_names:
                return clean_name, 0, 2
            
            # 3. Нечеткий поиск по известным именам
            extracted = process.extractOne(clean_name, self.known_names, scorer=fuzz.ratio)
            if extracted:
                best_match, score = extracted
                
                # score >= 75 И (разница длин < 7 ИЛИ начинается с ИЛИ score == 100)
                if score >= 75 and (
                    (abs(len(clean_name) - len(best_match)) < 7 and len(clean_name) > 2) or
                    best_match.startswith(clean_name) or
                    score == 100
                ):
                    return best_match, score, (0 if score == 100 else 1)
            
            # Регулярка прошла, но нечеткий поиск не нашел хорошего совпадения — это новое имя
            return clean_name, 0, 2
        
        # 4. Последний шанс / Очистка
        # Удаляем пунктуацию
        cleaned = re.sub(r"[.,_\-—\"]", "", raw_text)
        if len(cleaned) > 1:
            # Делаем первую букву заглавной
            cleaned = cleaned[0].upper() + cleaned[1:].lower()
            
            if self.known_names:
                extracted = process.extractOne(cleaned, self.known_names, scorer=fuzz.ratio)
                if extracted:
                    best_match, score = extracted
                    if len(cleaned) > 4 and score > 80:
                        return best_match, score, 1
            
            return cleaned, 0, 3 # Помечено как ошибка/неизвестно, но возвращено для ручной проверки
                
        return None, 0, 3
