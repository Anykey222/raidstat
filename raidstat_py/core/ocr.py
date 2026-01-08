import os
import re
import cv2
import numpy as np
import logging
import pytesseract
from PIL import Image
import sys
import requests
import io

class OCRHandler:
    def __init__(self, config=None, lang='ru'):
        self.logger = logging.getLogger(__name__)
        self.config = config
        self.lang = lang
        self._init_tesseract()

    def _init_tesseract(self):
        try:
            # Подавляем логгер pytesseract
            logging.getLogger('pytesseract').setLevel(logging.ERROR)
            
            # 1. Проверка встроенного Tesseract
            # При запуске в скомпилированном режиме (PyInstaller), ищем в sys._MEIPASS
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
                bundled_tesseract = os.path.join(base_path, 'bin', 'tesseract', 'tesseract.exe')
            else:
                # Нормальный режим разработки
                # raidstat_py/core/ocr.py -> raidstat_py -> bin
                current_dir = os.path.dirname(os.path.abspath(__file__))
                project_root = os.path.dirname(os.path.dirname(current_dir)) 
                bundled_tesseract = os.path.join(project_root, 'raidstat_py', 'bin', 'tesseract', 'tesseract.exe')
                
                # Запасной вариант, если структура отличается (например, запуск из корня)
                if not os.path.exists(bundled_tesseract):
                     bundled_tesseract = os.path.join(project_root, 'bin', 'tesseract', 'tesseract.exe')

            if os.path.exists(bundled_tesseract):
                pytesseract.pytesseract.tesseract_cmd = bundled_tesseract
                self.logger.info(f"Используется встроенный Tesseract: {bundled_tesseract}")
                return

            # 2. Проверка системного PATH
            try:
                pytesseract.get_tesseract_version()
                self.logger.info("Tesseract найден в системном PATH.")
                return
            except pytesseract.TesseractNotFoundError:
                pass
            
            self.logger.warning("Исполняемый файл Tesseract не найден. OCR не будет работать.")

        except Exception as e:
            self.logger.error(f"Не удалось инициализировать Tesseract: {e}")
            raise

    def recognize_text(self, image_path_or_array, crop_area=None, det=True, lang=None, config=None):
        """
        Распознает текст на изображении или в конкретной области кропа.
        
        Args:
            image_path_or_array: Путь к изображению или массив numpy (изображение cv2).
            crop_area: Кортеж (x, y, w, h) для кропа перед OCR.
            det: Использовать ли детектирование текста (установите False, если изображение уже является кропом текста).
            lang: Переопределить язык (например, 'eng', 'rus', 'eng+rus').
            config: Дополнительная строка конфигурации Tesseract.
            
        Returns:
            Список кортежей: [(текст, уверенность), ...]
        """
        img = self._load_image(image_path_or_array)
        if img is None: return []

        if crop_area:
            img = self._crop_image(img, crop_area)

        return self._recognize_tesseract(img, det, lang, config)

    def _load_image(self, image_path_or_array):
        if isinstance(image_path_or_array, str):
            return cv2.imread(image_path_or_array)
        elif isinstance(image_path_or_array, np.ndarray):
            return image_path_or_array
        else:
            self.logger.error("Некорректный ввод изображения. Должен быть путь или массив numpy.")
            return None

    def _crop_image(self, img, crop_area):
        x, y, w, h = crop_area
        h_img, w_img = img.shape[:2]
        x = max(0, min(x, w_img))
        y = max(0, min(y, h_img))
        w = max(1, min(w, w_img - x))
        h = max(1, min(h, h_img - y))
        return img[y:y+h, x:x+w]

    def preprocess_for_ocr(self, img, scale_factor=2, padding=5, use_otsu=True, invert=False, 
                           min_threshold=None, max_threshold=None, otsu_offset=0, fixed_threshold=None):
        """
        Предобработка изображения для улучшения качества распознавания OCR.
        
        Args:
            img: numpy array (cv2 изображение)
            scale_factor: коэффициент масштабирования (2-3 обычно хорошо работает для мелкого текста)
            padding: количество пикселей белого паддинга вокруг изображения
            use_otsu: использовать бинаризацию Otsu для отделения текста от фона
            invert: инвертировать цвета перед обработкой (для светлого текста на тёмном фоне)
            min_threshold: минимальная яркость пикселя (0-255). Пиксели темнее будут отброшены (стали белыми)
            max_threshold: максимальная яркость пикселя (0-255). Пиксели светлее будут отброшены (стали белыми)
            otsu_offset: смещение порога Otsu (-127 до +127). Положительное значение делает результат темнее
            fixed_threshold: использовать фиксированный порог вместо Otsu (0-255)
            
        Returns:
            Обработанное изображение в формате BGR (3 канала) для совместимости с OCR
        """
        if img is None:
            return None
            
        processed = img.copy()
        
        # Инверсия цветов если нужно (светлый текст на тёмном фоне -> тёмный на светлом)
        if invert:
            processed = cv2.bitwise_not(processed)
        
        # Конвертация в grayscale если ещё не сделано
        if len(processed.shape) == 3:
            gray = cv2.cvtColor(processed, cv2.COLOR_BGR2GRAY)
        else:
            gray = processed
        
        # Фильтрация по диапазону цветов (отбрасываем ненужные цвета)
        # Полезно для черного текста на сером фоне - можно отбросить слишком светлые области
        if min_threshold is not None or max_threshold is not None:
            filtered = gray.copy()
            
            # Отбрасываем пиксели темнее min_threshold (делаем их белыми)
            if min_threshold is not None:
                filtered[gray < min_threshold] = 255
            
            # Отбрасываем пиксели светлее max_threshold (делаем их белыми)
            if max_threshold is not None:
                filtered[gray > max_threshold] = 255
            
            gray = filtered
        
        # Масштабирование для улучшения распознавания мелкого текста
        if scale_factor and scale_factor > 1:
            h, w = gray.shape[:2]
            new_w = int(w * scale_factor)
            new_h = int(h * scale_factor)
            gray = cv2.resize(gray, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
        
        # Бинаризация
        if fixed_threshold is not None:
            # Фиксированный порог
            _, binary = cv2.threshold(gray, fixed_threshold, 255, cv2.THRESH_BINARY)
        elif use_otsu:
            # Бинаризация через Otsu's thresholding
            # Это помогает отделить текст от шумного фона
            otsu_threshold, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Применяем смещение к порогу Otsu, если указано
            if otsu_offset != 0:
                adjusted_threshold = max(0, min(255, otsu_threshold + otsu_offset))
                _, binary = cv2.threshold(gray, adjusted_threshold, 255, cv2.THRESH_BINARY)
        else:
            binary = gray
        
        # Добавляем белый паддинг вокруг изображения
        # Это помогает OCR лучше обнаруживать текст на краях
        if padding and padding > 0:
            padded = cv2.copyMakeBorder(
                binary, 
                padding, padding, padding, padding, 
                cv2.BORDER_CONSTANT, 
                value=255
            )
        else:
            padded = binary
        
        # Конвертация обратно в BGR для совместимости с OCR (Tesseract ожидает 3 канала)
        final = cv2.cvtColor(padded, cv2.COLOR_GRAY2BGR)
        
        return final

    def _recognize_tesseract(self, img, det, lang=None, extra_config=None):
        try:
            # Конвертируем в RGB для Pillow
            if len(img.shape) == 3:
                img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            else:
                img_rgb = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
                
            pil_img = Image.fromarray(img_rgb)
            
            # Выбор языка
            tess_lang = lang if lang else ('rus' if self.lang == 'ru' else 'eng')
            if self.lang == 'en' and not lang: tess_lang = 'eng'
            
            # Конфигурация
            # PSM 3: Полностью автоматическая сегментация страницы, но без OSD (по умолчанию).
            # PSM 6: Предполагается наличие одного однородного блока текста.
            # PSM 7: Изображение рассматривается как одна строка текста.
            
            # Если det=False, мы предполагаем, что это конкретный фрагмент (например, имя или число), поэтому PSM 7 обычно лучше.
            # Если det=True, это может быть большая область, поэтому PSM 3 или 6.
            
            psm_config = '--psm 7' if not det else '--psm 6'
            full_config = psm_config
            if extra_config:
                full_config += " " + extra_config
            full_config += " " + '--oem 1'
            
            data = pytesseract.image_to_data(pil_img, lang=tess_lang, config=full_config, output_type=pytesseract.Output.DICT)
            # if 'text' in data:
            #     for i in range(len(data['text'])):
            #         # Вывод соответствия текста и уверенности для отладки
            #         print(f"OCR Debug: '{data['text'][i]}' (conf: {data['conf'][i]})")
            
            extracted_data = []
            if 'text' in data:
                n_boxes = len(data['text'])
                for i in range(n_boxes):
                    # Отфильтровываем пустой текст и низкую уверенность
                    text = data['text'][i].strip()
                    conf = int(data['conf'][i])
                    
                    if text and conf >= 0:
                        # Нормализуем уверенность до 0.0 - 1.0
                        normalized_conf = float(conf) / 100.0
                        extracted_data.append((text, normalized_conf))
                        
            return extracted_data
        except Exception as e:
            self.logger.error(f"Ошибка распознавания Tesseract: {e}")
            return []

    def recognize_batch(self, image_list, det=True):
        """
        Распознает текст в списке изображений.
        Args:
            image_list: Список массивов numpy (изображений).
            det: Использовать ли детектирование.
        Returns:
            Список кортежей (текст, уверенность), соответствующих входным изображениям.
            Если текст для изображения не найден, возвращает (None, 0.0).
        """
        if not image_list:
            return []
            
        final_results = []
        
        for img in image_list:
            try:
                results = self.recognize_text(img, det=det)
                
                if results:
                    full_text = " ".join([item[0] for item in results])
                    avg_conf = sum([item[1] for item in results]) / len(results)
                    final_results.append((full_text.strip(), avg_conf))
                else:
                    final_results.append((None, 0.0))
            except Exception as e:
                self.logger.error(f"Ошибка при обработке изображения в пакете: {e}")
                final_results.append((None, 0.0))
                
        return final_results

    def recognize_single_line(self, image_path_or_array, crop_area=None, lang=None, config=None):
        """
        Помощник для получения одной строки из области (например, имя, число).
        Объединяет несколько обнаруженных блоков, если это необходимо.
        """
        # Для одной строки мы определенно хотим det=False (обычно PSM 7)
        results = self.recognize_text(image_path_or_array, crop_area, det=False, lang=lang, config=config)
        if not results:
            return None, 0.0
        
        full_text = " ".join([r[0] for r in results])
        avg_conf = sum([r[1] for r in results]) / len(results)
        
        return full_text.strip(), avg_conf

    def recognize_online_ocr_space(self, image_path_or_array, api_key=None, language='auto'):
        """
        Распознает текст с помощью OCR.space API.
        
        Args:
            image_path_or_array: Путь к изображению или массив numpy (изображение cv2).
            api_key: API ключ OCR.space. Если None, делается попытка взять из конфига.
            language: Код языка.
            
        Returns:
            Текст (строка) или None в случае ошибки.
        """
        try:
            if not api_key:
                if self.config:
                    api_key = self.config.get("ocr_api_key")
                
            if not api_key:
                self.logger.warning("Ключ API OCR.space не предоставлен и не найден в конфиге.")
                return None

            img = self._load_image(image_path_or_array)
            if img is None:
                return None

            # Конвертируем в байты
            is_success, buffer = cv2.imencode(".jpg", img)
            if not is_success:
                self.logger.error("Не удалось закодировать изображение для онлайн OCR")
                return None
            
            io_buf = io.BytesIO(buffer)
            
            payload = {
                'apikey': api_key,
                'language': language,
                'OCREngine': '2',
                'scale': 'true',
            }
            
            files = {
                'file': ('image.jpg', io_buf, 'image/jpeg')
            }
            
            # self.logger.debug("Отправка запроса к OCR.space API...")
            response = requests.post('https://api.ocr.space/parse/image',
                                     files=files,
                                     data=payload,
                                     timeout=10) # тайм-аут 10 секунд
            
            if response.status_code != 200:
                self.logger.error(f"OCR.space API HTTP Error: {response.status_code} - {response.text}")
                return None

            try:
                result = response.json()
            except ValueError:
                self.logger.error(f"API OCR.space вернул не-JSON ответ: {response.text}")
                return None
            
            if not isinstance(result, dict):
                self.logger.error(f"API OCR.space вернул неожиданный формат: {type(result)}")
                return None

            if result.get('IsErroredOnProcessing'):
                self.logger.error(f"OCR.space API error: {result.get('ErrorMessage')}")
                return None
                
            parsed_results = result.get('ParsedResults')
            if not parsed_results:
                return []
                
            # Extract text
            extracted_text = parsed_results[0].get('ParsedText')
            
            # OCR.space Engine 2 не всегда возвращает уверенность для каждого слова в простом выводе,
            # но мы можем вернуть текст.
            if extracted_text:
                return extracted_text.strip()
            
            return ""

        except Exception as e:
            self.logger.error(f"Ошибка запроса онлайн OCR: {e}")
            return None
    @staticmethod
    def _get_longest_word(text):
        if not text:
            return None
        words = text.split()
        if not words:
            return None
        return max(words, key=len)

    def process_name_recognition(self, full_img_bgr, rect, matcher, ocr_mode='offline', 
                                preprocess_params=None, online_crop_no_otsu=False, retry_with_shifts=True, item_id=""):
        """
        Унифицированный метод для распознавания имен с повторными попытками.
        
        Args:
            full_img_bgr: Полное изображение в BGR (или кроп, содержащий rect).
            rect: Кортеж (x, y, w, h), определяющий область кропа относительно full_img_bgr.
            matcher: Экземпляр Matcher для идентификации имен.
            ocr_mode: 'offline' или 'mixed'.
            preprocess_params: Словарь с 'use_otsu', 'max_threshold', 'otsu_offset' и т.д.
            online_crop_no_otsu: Если True, использует кроп без Otsu для онлайн-повтора.
            retry_with_shifts: Если True, пробует сдвиги y-1 и y+1, если результат не оптимален.
            item_id: Строковый ID для отладочных логов (например, координаты ячейки или имя файла).
            
        Returns:
            (name, score, type_code, crop_processed)
        """
        preprocess_params = preprocess_params or {}
        use_otsu = preprocess_params.get("use_otsu", True)
        max_threshold = preprocess_params.get("max_threshold", None)
        otsu_offset = preprocess_params.get("otsu_offset", 0) 
        padding = preprocess_params.get("padding", 0)
        
        debug_info = [] # Список для накопления шагов отладки
        
        x, y, w, h = rect
        # Проверка границ
        h_img, w_img = full_img_bgr.shape[:2]
        if y + h > h_img or x + w > w_img:
            return None, 0, 3, "OutOfBounds"

        crop_bgr = full_img_bgr[y:y+h, x:x+w].copy()
        
        # 1. Основная стратегия предобработки
        # Логика: если otsu_offset != 0, сначала пробуем ФИКСИРОВАННЫЙ порог (что помогает для «битых» ячеек).
        # Мы пробуем 75 как хорошую базовую линию для битых ячеек.
        current_fixed_threshold = 75 if (otsu_offset != 0) else None
        current_use_otsu = use_otsu if (otsu_offset == 0) else False # Если смещение задано, мы не используем Otsu на шаге 1, а используем фиксированный порог.
        
        crop_processed = self.preprocess_for_ocr(
            crop_bgr,
            scale_factor=2, 
            padding=padding, 
            use_otsu=current_use_otsu,
            max_threshold=max_threshold,
            otsu_offset=0, # Не используется здесь, если задан фиксированный порог
            fixed_threshold=current_fixed_threshold,
            invert=True
        )
        
        # 1.1 Сопоставление
        raw_text, conf = self.recognize_single_line(crop_processed, lang='eng+rus')
        longest_word = self._get_longest_word(raw_text)
        name, score, type_code = matcher.smart_match(longest_word)
        
        debug_info.append(f"[Step 1 Fixed otsu] Text='{raw_text}' Name='{name}' Type={type_code}")

        # 2. Повтор со смещением Otsu (если нужно)
        # Если не удалось (нет имени или мусор) И настроено смещение, пробуем с правильным Otsu + смещение.
        if (not name or type_code == 3) and otsu_offset != 0:
            crop_retry_otsu = self.preprocess_for_ocr(
                crop_bgr,
                scale_factor=2, 
                padding=padding, # Сохраняем паддинг
                use_otsu=True,
                max_threshold=max_threshold,
                otsu_offset=otsu_offset,
                fixed_threshold=None,
                invert=True
            )
            raw_text, conf = self.recognize_single_line(crop_retry_otsu, lang='eng+rus')
            longest_word = self._get_longest_word(raw_text)
            name_retry, score_retry, type_code_retry = matcher.smart_match(longest_word)
            
            debug_info.append(f"[Step 2 OtsuOffset] Text='{raw_text}' Name='{name_retry}' Type={type_code_retry}")
            
            if type_code_retry in [0, 1, 2, 4]:
                 name = name_retry
                 score = score_retry
                 type_code = type_code_retry
                 crop_processed = crop_retry_otsu # Обновляем валидный кроп

        # 3. Повтор без Otsu (чистая версия)
        if (not name or type_code == 3):
             crop_no_otsu = self.preprocess_for_ocr(
                crop_bgr,
                scale_factor=2, 
                padding=0, 
                use_otsu=False,
                max_threshold=0,
                otsu_offset=0,
                invert=True
             )
             raw_text, conf = self.recognize_single_line(crop_no_otsu, lang='eng+rus')
             longest_word = self._get_longest_word(raw_text)
             name_retry, score_retry, type_code_retry = matcher.smart_match(longest_word)
             
             debug_info.append(f"[Шаг 3 БезOtsu]: Текст='{raw_text}' Имя='{name_retry}' Тип={type_code_retry}")

             if type_code_retry in [0, 1, 2, 4]:
                  name = name_retry
                  score = score_retry
                  type_code = type_code_retry
                  crop_processed = crop_no_otsu

        # 3.5. Повтор со сменой языка
        # Если оффлайн режим и результат типа 2 (Новый) или 3 (Мусор), пробуем конкретный язык
        if ocr_mode == 'offline' and type_code in [2, 3] and name:
            has_cyrillic = bool(re.search('[а-яА-ЯёЁ]', name))
            target_lang = 'eng' if has_cyrillic else 'rus'
            
            raw_text_retry, conf_retry = self.recognize_single_line(crop_processed, lang=target_lang)
            longest_word_retry = self._get_longest_word(raw_text_retry)
            name_retry, score_retry, type_code_retry = matcher.smart_match(longest_word_retry)
            
            debug_info.append(f"[Step 3.5 LangRetry] {target_lang}: Text='{raw_text_retry}' Name='{name_retry}' Type={type_code_retry}")

            if type_code_retry in [0, 1, 4]:
                 name = name_retry
                 score = score_retry
                 type_code = type_code_retry

        # 4. Смешанный режим (Online)
        if ocr_mode == 'mixed' and (not name or type_code in [2, 3]):
            try:
                if online_crop_no_otsu:
                    crop_online = self.preprocess_for_ocr(
                        crop_bgr,
                        scale_factor=2, 
                        padding=0, 
                        use_otsu=False,
                        max_threshold=max_threshold,
                        otsu_offset=otsu_offset,
                        invert=True
                    )
                else:
                    crop_online = crop_processed
                    
                online_text = self.recognize_online_ocr_space(crop_online)
                debug_info.append(f"[Step 4 Online]: Text='{online_text}'")
                
                if online_text:
                    longest_word_online = self._get_longest_word(online_text)
                    name_online, score_online, type_code_online = matcher.smart_match(longest_word_online)
                    
                    if name_online:
                        debug_info.append(f" -> Имя='{name_online}' Тип={type_code_online}")
                        name = name_online
                        score = score_online
                        type_code = type_code_online
            except Exception as e:
                self.logger.error(f"Ошибка онлайн OCR: {e}")
                debug_info.append(f"[Step 4 Online]: Error {e}")

        # 5. Повтор со сдвигом области (новое)
        if retry_with_shifts and (not name or type_code in [2, 3]):
            # Помощник для сдвига
            def try_shift(shift_pix):
                new_y = y + shift_pix
                if new_y < 0: return None
                
                # Рекурсия
                res = self.process_name_recognition(
                    full_img_bgr, (x, new_y, w, h), matcher, 'offline', 
                    preprocess_params, online_crop_no_otsu, retry_with_shifts=False, item_id=item_id + "_shift"
                )
                return res
 
            # Пробуем ВВЕРХ
            res_up = try_shift(-1)
            if res_up:
                n_up, s_up, tc_up, _ = res_up # игнорируем возврат кропа
                debug_info.append(f"[Step 5 Shift -1]: Name='{n_up}' Type={tc_up}")
                if tc_up in [0, 1]:
                    name = n_up
                    score = s_up
                    type_code = tc_up
            
            if type_code not in [0, 1]:
                # Пробуем ВНИЗ
                res_down = try_shift(1)
                if res_down:
                    n_down, s_down, tc_down, _ = res_down
                    debug_info.append(f"[Step 5 Shift +1]: Name='{n_down}' Type={tc_down}")
                    if tc_down in [0, 1]:
                        name = n_down
                        score = s_down
                        type_code = tc_down
            
            # Если мы здесь, сдвиги не помогли (или не были идеальными). Возвращаем оригинал.
            
        if not item_id.endswith("_shift"):
             is_debug = False
             if self.config:
                 # Проверяем, является ли конфиг объектом с атрибутом debug или словарем
                 if hasattr(self.config, 'debug'):
                     is_debug = self.config.debug
                 elif isinstance(self.config, dict):
                     is_debug = self.config.get('debug', False)
             
             if is_debug:
                 full_log = f"[{item_id}] " + " | ".join(debug_info)
                 self.logger.info(full_log)
             else:
                 self.logger.info(f"[{item_id}] -> '{name}'")

        return name, score, type_code, crop_processed
