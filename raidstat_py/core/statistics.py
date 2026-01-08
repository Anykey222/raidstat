import os
import logging
from datetime import datetime
import numpy as np
import cv2
from PIL import Image
from .ocr import OCRHandler
from .matcher import Matcher
from ..utils.config import Config
import concurrent.futures
import shutil
from .history import HistoryManager

class StatisticsProcessor:
    def __init__(self, config: Config, ocr: OCRHandler, matcher: Matcher, storage, debug_screens=False):
        self.config = config
        self.ocr = ocr
        self.matcher = matcher
        self.storage = storage
        self.logger = logging.getLogger(__name__)
        self.debug_screens = debug_screens
        
        self.scale = self.config.interface_scale
        self.offsets = self._get_offsets(self.scale)
        self.history = HistoryManager("statistics")

    def revert_history(self):
        self.history.revert()

    def _get_offsets(self, scale):
        # Перенесено из RaidStat.java
        start_x = self.config.personal_frame_coords['x']
        start_y = self.config.personal_frame_coords['y']
 
        offsets = {}
        
        # Смещения относительно start_x, start_y
        if scale == 130: # 130%
            offsets['name'] = (-327, -20, 140, 19)
            offsets['class'] = (57, 10, 367, 22)
            offsets['honor'] = (241, 125, 89, 16)
            offsets['kills'] = (207, 148, 79, 16)
            offsets['gear'] = (179, 79, 63, 16)
        elif scale == 120:
            offsets['name'] = (-302, -20, 118, 18)
            offsets['class'] = (53, 9, 362, 20)
            offsets['honor'] = (225, 113, 80, 15)
            offsets['kills'] = (195, 133, 70, 15)
            offsets['gear'] = (167, 72, 58, 15)
        elif scale == 110:
            offsets['name'] = (-276, -18, 95, 17)
            offsets['class'] = (49, 8, 357, 18)
            offsets['honor'] = (210, 105, 71, 15)
            offsets['kills'] = (180, 124, 61, 15)
            offsets['gear'] = (155, 67, 53, 15)
        else: # 100
            offsets['name'] = (-251, -15, 85, 15)
            offsets['class'] = (47, 8, 352, 17)
            offsets['honor'] = (193, 98, 62, 14)
            offsets['kills'] = (165, 116, 52, 14)
            offsets['gear'] = (143, 63, 48, 14)
            
        return offsets

    def process_folder(self, folder_path, recursive=False, stop_event=None):
        self.history.clear()
        image_files = []
        for root, dirs, files in os.walk(folder_path):
            if stop_event and stop_event.is_set():
                self.logger.info("Обработка статистики прервана (этап сканирования).")
                return 0

            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                    image_files.append(os.path.join(root, file))
            if not recursive:
                break
        
        if not image_files:
            return 0
            
        # Сортировка по времени модификации
        image_files.sort(key=os.path.getmtime)
        
        # Группировка по времени
        groups = []
        if not image_files:
            return 0
            
        current_group = [image_files[0]]
        group_start_time = os.path.getmtime(image_files[0])
        max_diff = self.config.max_diff_time * 60 # секунды
        
        for img_path in image_files[1:]:
            curr_time = os.path.getmtime(img_path)
            if curr_time - group_start_time > max_diff:
                # Сохраняем текущую группу и начинаем новую
                groups.append(current_group)
                current_group = [img_path]
                group_start_time = curr_time
            else:
                current_group.append(img_path)
        if current_group:
            groups.append(current_group)
            
        # Обработка каждой группы
        total_processed = 0
        first_group = True
        prev_group_stats = None
        
        try:
            for group in groups:
                if stop_event and stop_event.is_set():
                    self.logger.info("Обработка статистики прервана.")
                    return total_processed

                self.logger.info(f"Обработка группы с {len(group)} изображениями")
                
                # Получаем дату/время первого файла в группе для создания подпапки
                first_file_time = os.path.getmtime(group[0])
                date_str = datetime.fromtimestamp(first_file_time).strftime("%Y-%m-%d")
                time_str = datetime.fromtimestamp(first_file_time).strftime("%H-%M")
                
                # Создаем подпапку для группы
                group_folder = os.path.join(folder_path, date_str, time_str)
                if not os.path.exists(group_folder):
                    os.makedirs(group_folder)
                    self.history.add_created(group_folder)
                else:
                    os.makedirs(group_folder, exist_ok=True)
                
                group_stats, failed_paths = self.process_group(group, stop_event=stop_event)

                # Если отменили внутри групповой обработки
                if stop_event and stop_event.is_set():
                    self.logger.info("Обработка статистики прервана во время обработки группы.")
                    return total_processed
                
                # Логика как в Java: пропускаем сохранение первой группы
                # ВАЖНО: сохраняем статистику ДО перемещения файлов, пока debug_images доступны
                if not first_group and prev_group_stats:
                    # Обновляем текущую статистику на основе предыдущей группы
                    self.update_stats_between_groups(group_stats, prev_group_stats)
                    
                    # Сохраняем статистику (добавляем новое событие)
                    self.storage.save_statistics(group_stats, f"{date_str} {time_str.replace('-', ':')}", debug_screens=self.debug_screens)
                    total_processed += len(group) - len(failed_paths)
                
                # Создаем папку для ошибок если есть неудачные файлы
                errors_folder = os.path.join(folder_path, "errors")
                if failed_paths:
                    if not os.path.exists(errors_folder):
                        os.makedirs(errors_folder)
                        self.history.add_created(errors_folder)
                    else:
                        os.makedirs(errors_folder, exist_ok=True)

                # Перемещаем обработанные файлы
                for img_path in group:
                    if stop_event and stop_event.is_set():
                        break # Не перемещаем, если прервано прямо здесь
                    try:
                        filename = os.path.basename(img_path)
                        
                        if img_path in failed_paths:
                            dest_path = os.path.join(errors_folder, filename)
                            self.logger.warning(f"Moving failed file {filename} to errors folder")
                        else:
                            dest_path = os.path.join(group_folder, filename)
                        
                        if os.path.exists(dest_path):
                            os.remove(dest_path)
                        shutil.move(img_path, dest_path)
                        self.history.add_move(img_path, dest_path)
                        
                        # Перемещаем папку отладки, если она существует
                        if self.debug_screens:
                            debug_dir_name = os.path.splitext(filename)[0]
                            src_debug_dir = os.path.join(folder_path, debug_dir_name)
                            if os.path.exists(src_debug_dir):
                                # Если неудача, возможно, тоже стоит оставить в ошибках?
                                if img_path in failed_paths:
                                    dest_debug_dir = os.path.join(errors_folder, debug_dir_name)
                                else:
                                    dest_debug_dir = os.path.join(group_folder, debug_dir_name)
                                    
                                if os.path.exists(dest_debug_dir):
                                    shutil.rmtree(dest_debug_dir)
                                shutil.move(src_debug_dir, dest_debug_dir)
                                self.history.add_created(dest_debug_dir) # Отслеживаем как созданное/перемещенное, чтобы можно было откатить
                                self.history.add_move(src_debug_dir, dest_debug_dir)
 
                    except Exception as e:
                        self.logger.error(f"Ошибка при перемещении файла {img_path}: {e}")
                
                # Обновляем пути к debug_images после перемещения файлов
                # Это важно для того, чтобы скриншоты "до" были доступны при обработке следующей группы
                if self.debug_screens:
                    self._update_debug_paths_after_move(group_stats, folder_path, group_folder, errors_folder, failed_paths)
                
                # Сохраняем текущую группу для следующей итерации
                prev_group_stats = group_stats
                first_group = False
        finally:
            self.history.save()
            
        return total_processed

    def process_group(self, image_paths, stop_event=None):
        """
        Обработка группы изображений, представляющих одно событие.
        Логика:
        - Идентификация уникальных лиц.
        - Для каждого лица поиск начальных характеристик (первое появление) и конечных характеристик (последнее появление).
        - Вычисление Дельты = Конец - Начало.
        """
        person_data = {} # {name: {start: {}, end: {}}}
        
        num_threads = 8
        
        self.logger.debug(f"Обработка группы в {num_threads} потоков")
        
        # Потокобезопасный set для отслеживания дубликатов (как namesPerDate в Java)
        import threading
        names_lock = threading.Lock()
        names_per_group = {}
        
        def safe_process_image(path):
            if stop_event and stop_event.is_set():
                return None
            try:
                return self.process_image(path, names_per_group, names_lock, stop_event=stop_event)
            except Exception as e:
                self.logger.error(f"Не удалось обработать {path}: {e}")
                return None

        results = []
        if num_threads > 1:
            with concurrent.futures.ThreadPoolExecutor(max_workers=num_threads) as executor:
                futures = [executor.submit(safe_process_image, path) for path in image_paths]
                # Важно: итерируемся по futures в исходном порядке, а не через as_completed,
                # чтобы результаты соответствовали порядку файлов в image_paths.
                for future in futures:
                    if stop_event and stop_event.is_set():
                         # cancel rest
                         for f in futures:
                             f.cancel()
                         self.logger.info("Обработка группы статистики прервана.")
                         break
                    results.append(future.result())
        else:
            for path in image_paths:
                if stop_event and stop_event.is_set():
                    break
                results.append(safe_process_image(path))
        
        if stop_event and stop_event.is_set():
            return {}, []
        
        failed_paths = []
        
        for path, stats in zip(image_paths, results):
            # Дубликат — пропускаем (но не добавляем в failed_paths!)
            if stats and stats.get('duplicate'):
                continue
            # None или нет имени — это ошибка
            if not stats or not stats.get('name'):
                failed_paths.append(path)
                continue
                
            name = stats['name']
            
            # При новой логике повторов (для исправления плохих данных), мы должны перезаписывать и 'start',
            # так как предыдущий 'start' был "плохим".
            # Если дубликаты отключены или это "исправляющий" дубликат, мы обновляем запись целиком.
            person_data[name] = {'start': stats, 'end': stats}
                
        # Подготовка финальных данных
        final_stats = {}
        for name, data in person_data.items():
            start = data['start']
            end = data['end']
            
            # Сохраняем None как есть (не заменяем на 0)
            kills_start = start.get('kills')
            kills_end = end.get('kills')
            honor_start = start.get('honor')
            honor_end = end.get('honor')
            
            # Дельта внутри группы — только если оба значения есть
            if kills_start is not None and kills_end is not None:
                kills_delta = kills_end - kills_start
            else:
                kills_delta = None
                
            if honor_start is not None and honor_end is not None:
                honor_delta = honor_end - honor_start
            else:
                honor_delta = None
            
            final_stats[name] = {
                'class': end.get('class'),
                'gear': end.get('gear'),
                'kills': kills_delta,  # Дельта (будет пересчитана между группами)
                'honor': honor_delta,  # Дельта (будет пересчитана между группами)
                'kills_start': kills_start,  # Начало группы
                'kills_end': kills_end,      # Конец группы
                'honor_start': honor_start,  # Начало группы
                'honor_end': honor_end,      # Конец группы
                'debug_images_start': start.get('debug_images'),  # Скриншоты начала (для "до")
                'debug_images_end': end.get('debug_images')       # Скриншоты конца (для "после")
            }
            
        return final_stats, failed_paths
    
    def update_stats_between_groups(self, current_stats, prev_stats):
        """
        Обновляет статистику текущей группы на основе предыдущей.
        Логика: конец предыдущей группы становится началом текущей.
        Также добавляет имена, которые были в предыдущей группе, но отсутствуют в текущей (вылет/уход),
        сохраняя их "до" значения и выставляя 0 прогресса.
        """
        # 1. Обновляем тех, кто есть в текущей группе
        for name, data in current_stats.items():
            if name in prev_stats:
                # Используем конечные значения предыдущей группы как начальные текущей
                old_honor_end = prev_stats[name].get('honor_end')
                old_kills_end = prev_stats[name].get('kills_end')
                
                # Обновляем начальные значения
                data['honor_start'] = old_honor_end
                data['kills_start'] = old_kills_end
                
                if old_honor_end is not None and old_kills_end is not None:
                    # Пересчитываем дельты
                    new_honor_delta = data['honor_end'] - old_honor_end
                    new_kills_delta = data['kills_end'] - old_kills_end
                    
                    data['honor'] = new_honor_delta
                    data['kills'] = new_kills_delta
                else:
                    # Если предыдущее значение неизвестно (например, человек отсутствовал),
                    # то дельту посчитать нельзя.
                    data['honor'] = None
                    data['kills'] = None
                
                # Скриншоты "до" берём из конца предыдущей группы
                data['debug_images_start'] = prev_stats[name].get('debug_images_end')
            else:
                # Если имени не было в предыдущей группе, мы не знаем начальных значений для этого события.
                # Поэтому дельты и начальные значения должны быть пустыми.
                data['honor'] = None
                data['kills'] = None
                data['honor_start'] = None
                data['kills_start'] = None
                data['debug_images_start'] = None

        # 2. Добавляем тех, кто был в предыдущей группе, но нет в текущей (ушли/вылетели)
        for name, prev_data in prev_stats.items():
            if name not in current_stats:
                honor_val = prev_data['honor_end']
                kills_val = prev_data['kills_end']
                
                current_stats[name] = {
                    'class': prev_data['class'],
                    'gear': prev_data['gear'],
                    'kills': None, # Пусто
                    'honor': None, # Пусто
                    'kills_start': kills_val,
                    'kills_end': None, # Пусто
                    'honor_start': honor_val,
                    'honor_end': None, # Пусто
                    'debug_images_start': prev_data.get('debug_images_end'), # Скриншот "до"
                    'debug_images_end': None # Скриншота "после" нет
                }
    
    def _update_debug_paths_after_move(self, group_stats, folder_path, group_folder, errors_folder, failed_paths):
        """
        Обновляет пути к debug_images после перемещения файлов.
        Это необходимо для того, чтобы скриншоты "до" были доступны при обработке следующей группы.
        """
        for name, data in group_stats.items():
            for images_key in ['debug_images_start', 'debug_images_end']:
                images = data.get(images_key)
                if not images:
                    continue
                
                updated_images = {}
                for field, old_path in images.items():
                    if not old_path:
                        updated_images[field] = None
                        continue
                    
                    # Определяем новый путь на основе того, куда был перемещён скриншот
                    filename = os.path.basename(old_path)
                    # Путь к изображению обычно: folder_path/ScreenShotXXXX/field.jpg
                    # После перемещения: group_folder/ScreenShotXXXX/field.jpg
                    parent_dir = os.path.basename(os.path.dirname(old_path))  # ScreenShotXXXX
                    
                    # Проверяем, был ли файл в failed_paths (определяем по имени папки debug)
                    # Имя папки debug = имя скриншота без расширения
                    screenshot_name = parent_dir  # например "ScreenShot0001"
                    
                    # Ищем соответствующий путь скриншота в failed_paths
                    is_failed = any(screenshot_name in fp for fp in failed_paths)
                    
                    if is_failed:
                        new_path = os.path.join(errors_folder, parent_dir, filename)
                    else:
                        new_path = os.path.join(group_folder, parent_dir, filename)
                    
                    if os.path.exists(new_path):
                        updated_images[field] = new_path
                    else:
                        # Если новый путь не существует, оставляем старый (может ещё не перемещён)
                        updated_images[field] = old_path if os.path.exists(old_path) else None
                data[images_key] = updated_images

    def process_image(self, image_path, names_per_group=None, names_lock=None, stop_event=None):
        if stop_event and stop_event.is_set():
            return None

        filename = os.path.basename(image_path)

        img = Image.open(image_path)
        img_np = np.array(img)
        # Convert to BGR for consistent OCR/OpenCV handling
        img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        
        start_x = self.config.personal_frame_coords['x']
        start_y = self.config.personal_frame_coords['y']
        
        results = {}
        
        # Помощник для обрезки
        def get_crop(field_name):
            ox, oy, w, h = self.offsets[field_name]
            x = start_x + ox
            y = start_y + oy
            
            h_img, w_img = img_bgr.shape[:2]
            x = max(0, min(x, w_img))
            y = max(0, min(y, h_img))
            w = max(1, min(w, w_img - x))
            h = max(1, min(h, h_img - y))
            
            return img_bgr[y:y+h, x:x+w]

        # 1. Обработка имени
        ox, oy, w, h = self.offsets['name']
        x = start_x + ox
        y = start_y + oy
        rect = (x, y, w, h)
        
        # Специфические параметры предобработки для статистики
        stats_preprocess = {
            "use_otsu": True,
            "padding": 3,
            # Инверсия обрабатывается process_name_recognition, но мы можем передать её, если захотим кастомную
        }

        if stop_event and stop_event.is_set():
            return None

        name_val, score, type_code, name_crop = self.ocr.process_name_recognition(
            img_bgr,
            rect,
            self.matcher,
            ocr_mode=getattr(self.config, 'ocr_mode', 'offline'),
            preprocess_params=stats_preprocess,
            online_crop_no_otsu=False,
            retry_with_shifts=True,
            item_id=filename
        )

        if not name_val or type_code == 3:
             if self.debug_screens:
                 text_debug = "Failed" # Упрощенная отладочная информация
                 self.logger.info(f"Пропуск {filename} - имя не совпало")
                 # Сохраняем отладочное изображение для несовпавшего имени
                 debug_dir = os.path.join(os.path.dirname(image_path), os.path.splitext(filename)[0])
                 if not os.path.exists(debug_dir):
                     os.makedirs(debug_dir)
                     self.history.add_created(debug_dir)
                 else:
                    os.makedirs(debug_dir, exist_ok=True)
                 if name_crop is not None:
                    Image.fromarray(name_crop).save(os.path.join(debug_dir, "name_failed.jpg"))
             return None
        
        # Проверка на дубликат с учетом качества распознавания
        # Если имя уже есть, но предыдущий скан был без фрагов/хонора, пробуем перезаписать
        should_skip = False
        if names_per_group is not None and names_lock is not None:
            import time as time_module
            
            # Ждём, пока другой поток закончит обработку этого имени (если он её начал)
            max_wait_time = 5.0  # Максимальное время ожидания в секундах
            wait_interval = 0.05  # Интервал проверки
            waited_time = 0
            
            while True:
                if stop_event and stop_event.is_set():
                    return None
                
                with names_lock:
                    if name_val not in names_per_group:
                        # Первое появление имени — регистрируем и продолжаем
                        names_per_group[name_val] = {'valid': False, 'processing': True}
                        break
                    
                    prev_info = names_per_group[name_val]
                    
                    # Если ранее уже успешно считали (valid=True), то это дубль
                    if prev_info.get('valid', False):
                        should_skip = True
                        break
                    
                    # Если другой поток ещё обрабатывает — ждём
                    if prev_info.get('processing', False):
                        if waited_time >= max_wait_time:
                            # Таймаут — пропускаем
                            self.logger.debug(f"{filename} {name_val} (таймаут ожидания параллельной обработки)")
                            should_skip = True
                            break
                        # Продолжаем ожидание (освобождаем lock и спим)
                    else:
                        # Предыдущая обработка завершилась невалидно — пробуем сами
                        self.logger.info(f"{filename} {name_val} (повтор для уточнения данных)")
                        prev_info['processing'] = True
                        break
                
                # Если дошли сюда — нужно подождать
                time_module.sleep(wait_interval)
                waited_time += wait_interval

        if should_skip:
            self.logger.info(f"{filename} {name_val} (дубль)")
            return {'duplicate': True, 'name': name_val}
             
        results['name'] = name_val
        
        if stop_event and stop_event.is_set():
            return None

        # 2. Обработка класса
        class_crop_raw = get_crop('class')
        # Предобработка: инверсия, бинаризация, паддинг
        class_crop = self.ocr.preprocess_for_ocr(class_crop_raw, scale_factor=2, padding=5, use_otsu=False, invert=True)
        text, conf = self.ocr.recognize_single_line(class_crop, lang='rus')
        if text:
            # Очистка обратной кавычки и извлечение имени класса перед скобкой
            results['class'] = text.replace('`', '').replace("'", '').split('(')[0].strip()
        else:
            results['class'] = None
            
        # 3. Числовые поля
        # Для числовых полей храним обработанные изображения для отладки
        numeric_crops = {}
        has_valid_stats = False # Флаг: считались ли фраги или хонор
        
        for field in ['kills', 'honor', 'gear']:
            if stop_event and stop_event.is_set():
                return None
 
            crop_raw = get_crop(field)
            # Предобработка для чисел: инверсия, бинаризация, паддинг (без масштабирования)
            crop = self.ocr.preprocess_for_ocr(crop_raw, scale_factor=2, padding=5, use_otsu=False, invert=True)
            numeric_crops[field] = crop
            
            # Используем белый список цифр
            text, conf = self.ocr.recognize_single_line(crop, lang='eng', config='-c tessedit_char_whitelist=0123456789')
            
            if text:
                digits = "".join(filter(str.isdigit, text))
                if digits:
                    results[field] = int(digits)
                else:
                    results[field] = None  # Нет цифр — None вместо 0
            else:
                results[field] = None  # Нет текста — None вместо 0

        # Статы валидны только если распознаны и фраги, и хонор
        has_valid_stats = results.get('kills') is not None and results.get('honor') is not None

        # Обновляем статус валидности в общем словаре
        if names_per_group is not None and names_lock is not None:
            with names_lock:
                # Если текущий результат валиден (есть статы), обновляем запись
                if has_valid_stats:
                    names_per_group[name_val] = {'valid': True, 'processing': False}
                
                # Если мы не получили валидных данных, а в базе уже есть валидные (от другого потока),
                # то считаем текущий результат дублем/мусором, чтобы не перезатереть хорошее.
                elif names_per_group.get(name_val, {}).get('valid', False):
                     return {'duplicate': True, 'name': name_val}
                else:
                    # Если не удалось получить статы, снимаем флаг обработки, чтобы другие потоки могли попытаться
                    if name_val in names_per_group:
                        names_per_group[name_val]['processing'] = False
                
        
            self.logger.info(f"{filename}: {results}")
            
            # Сохранение отладочных изображений
            if self.debug_screens:
                debug_dir = os.path.join(os.path.dirname(image_path), os.path.splitext(filename)[0])
                if not os.path.exists(debug_dir):
                    os.makedirs(debug_dir)
                    self.history.add_created(debug_dir)
                else:
                    os.makedirs(debug_dir, exist_ok=True)
                
                debug_images = {}
                try:
                    name_path = os.path.join(debug_dir, "name.jpg")
                    Image.fromarray(name_crop).save(name_path)
                    debug_images['name'] = name_path
                    
                    class_path = os.path.join(debug_dir, "class.jpg")
                    Image.fromarray(class_crop).save(class_path)
                    debug_images['class'] = class_path
                    
                    for field in ['kills', 'honor', 'gear']:
                         field_path = os.path.join(debug_dir, f"{field}.jpg")
                         Image.fromarray(numeric_crops[field]).save(field_path)
                         debug_images[field] = field_path
                         
                    results['debug_images'] = debug_images
                except Exception as e:
                    self.logger.error(f"Не удалось сохранить отладочные изображения: {e}")
            
        return results
