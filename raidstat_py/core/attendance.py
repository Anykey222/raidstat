import os
import re
from PIL import Image, ImageDraw, ImageFont
import logging
from datetime import datetime
from .ocr import OCRHandler
from .matcher import Matcher
from ..utils.config import Config
import concurrent.futures
import shutil
from .history import HistoryManager
import cv2
import cv2
import numpy as np

ATTENDANCE_PREPROCESS = {
    "use_otsu": True,
    "max_threshold": None,
    "otsu_offset": -8, # Жестко заданное смещение для посещаемости
    "padding": 3
}

class AttendanceProcessor:
    def __init__(self, config: Config, ocr: OCRHandler, matcher: Matcher, storage):
        self.config = config
        self.ocr = ocr
        self.matcher = matcher
        self.storage = storage
        self.logger = logging.getLogger(__name__)
        
        # Настройки сетки на основе масштаба интерфейса
        self.scale = self.config.interface_scale
        self.grid_params = self._get_grid_params(self.scale)
        self.history = HistoryManager("attendance")

    def revert_history(self):
        self.history.revert()

    def _get_grid_params(self, scale):
        params = {}
        start_x = self.config.raid_frame_coords['x']
        start_y = self.config.raid_frame_coords['y']
        
        if scale == 130:
            start_x += 2
            start_y += 5
            params['cols_x'] = [
                start_x, 
                start_x + 88, 
                start_x + 88*2, 
                start_x + 88*3, 
                start_x + 88*4
            ]
            params['rows_y'] = [
                start_y,
                start_y + 44,
                start_y + 44*2,
                start_y + 44*3,
                start_y + 44*4
            ]
            params['shift_y'] = 264 
            params['shift_text_y'] = 20 
            params['name_w'] = 75
            params['name_h'] = 16
            params['font_size'] = 13
            
        elif scale == 120:
            start_x += 2
            start_y += 4
            params['cols_x'] = [
                start_x, 
                start_x + 81, 
                start_x + 81*2, 
                start_x + 81*2 + 80, 
                start_x + 81*2 + 80 + 81
            ]
            params['rows_y'] = [
                start_y,
                start_y + 41,
                start_y + 41 + 40,
                start_y + 41 + 40 + 41,
                start_y + 41 + 40 + 41 + 41
            ]
            params['shift_y'] = 245
            params['shift_text_y'] = 19
            params['name_w'] = 69
            params['name_h'] = 16
            params['font_size'] = 12
            
        elif scale == 110:
            start_x += 2
            start_y += 3
            params['cols_x'] = [
                start_x, 
                start_x + 74, 
                start_x + 74*2, 
                start_x + 74*3, 
                start_x + 74*4
            ]
            params['rows_y'] = [
                start_y,
                start_y + 37,
                start_y + 37 + 38,
                start_y + 37 + 38 + 37,
                start_y + 37 + 38 + 37 + 38
            ]
            params['shift_y'] = 224
            params['shift_text_y'] = 17
            params['name_w'] = 62
            params['name_h'] = 16
            params['font_size'] = 12
            
        else: # 100
            start_x += 2
            start_y += 2
            params['cols_x'] = [
                start_x, 
                start_x + 67, 
                start_x + 67 + 68, 
                start_x + 67 + 68 + 67, 
                start_x + 67 + 68 + 67 + 67
            ]
            params['rows_y'] = [
                start_y,
                start_y + 34,
                start_y + 34+35,
                start_y + 34+35+34,
                start_y + 34+35+34+35
            ]
            params['shift_y'] = 204
            params['shift_text_y'] = 14
            params['name_w'] = 52
            params['name_h'] = 15
            params['font_size'] = 11
            
        return params

    def process_folder(self, folder_path, recursive=False, stop_event=None):
        self.history.clear()
        
        # 1. Сбор файлов, сгруппированных по директориям
        # Структура: { путь_к_директории: [пути_к_файлам] }
        grouped_files = {}
        
        for root, dirs, files in os.walk(folder_path):
            if stop_event and stop_event.is_set():
                self.logger.info("Обработка посещаемости прервана (фаза сканирования).")
                return 0

            # Вычисляем глубину относительно folder_path
            rel_path = os.path.relpath(root, folder_path)
            if rel_path == '.':
                depth = 0
            else:
                depth = rel_path.count(os.sep) + 1
            
            # Логика:
            # Если recursive=False: только глубина 0 (корень)
            # Если recursive=True: глубина 0 и 1 (корень + подпапки первого уровня)
            if not recursive and depth > 0:
                continue
            if recursive and depth > 1:
                # Пропускаем подпапки более низких уровней
                continue
                
            current_files = []
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                    current_files.append(os.path.join(root, file))
            
            if current_files:
                # Сортируем по времени создания внутри группы
                current_files.sort(key=os.path.getmtime)
                grouped_files[root] = current_files
                
            # Изменяем dirs для обрезки os.walk
            if not recursive:
                dirs.clear() # Перестаем спускаться вниз
            elif depth >= 1:
                dirs.clear() # Перестаем спускаться глубже уровня 1

        if not grouped_files:
            return 0

        total_unique = 0
        num_threads = 8
        self.logger.info(f"Обработка посещаемости в {num_threads} потоков, найдено групп: {len(grouped_files)}")

        try:
            # Сортируем группы по времени изменения первого файла
            sorted_groups = sorted(grouped_files.items(), key=lambda item: os.path.getmtime(item[1][0]) if item[1] else 0)

            for group_path, image_files in sorted_groups:
                if stop_event and stop_event.is_set():
                    break
                    
                group_attendees = set()
                
                # Определяем название колонки
                if os.path.abspath(group_path) == os.path.abspath(folder_path):
                    # Корневая папка — используем временную метку первого файла
                    if image_files:
                         ts = os.path.getmtime(image_files[0])
                         column_name = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
                    else:
                        column_name = datetime.now().strftime("%Y-%m-%d %H:%M")
                else:
                    # Подпапка — используем имя папки
                    column_name = os.path.basename(group_path)
                
                self.logger.info(f"Обработка группы: {column_name} ({len(image_files)} изображений)")

                def safe_process_image(path):
                    if stop_event and stop_event.is_set():
                        return []
                    try:
                        return self.process_image(path, stop_event=stop_event)
                    except Exception as e:
                        self.logger.error(f"Не удалось обработать {path}: {e}")
                        return []

                results = []
                if num_threads > 1:
                     with concurrent.futures.ThreadPoolExecutor(max_workers=num_threads) as executor:
                        futures = [executor.submit(safe_process_image, path) for path in image_files]
                        for future in concurrent.futures.as_completed(futures):
                            if stop_event and stop_event.is_set():
                                for f in futures: f.cancel()
                                break
                            results.append(future.result())
                else:
                    for path in image_files:
                        if stop_event and stop_event.is_set():
                            break
                        results.append(safe_process_image(path))

                for attendees in results:
                    if attendees:
                        group_attendees.update(attendees)

                if group_attendees:
                    self.storage.save_attendance(list(group_attendees), column_name)
                    total_unique += len(group_attendees)
            
            if stop_event and stop_event.is_set():
                 return 0
                 
            return total_unique
        finally:
            self.history.save()

    def _process_single_cell(self, img_bgr, block_idx, row_idx, col_idx, x, curr_y, w, h, debug_dir):
        # Унифицированный вызов распознавания
        rect = (x, curr_y, w, h)
        
        # Логика параметров для отладочной ячейки (самой первой)
        if block_idx == 0 and row_idx == 0 and col_idx == 0:
             # self.logger.info(f"Параметры предобработки: {ATTENDANCE_PREPROCESS}")
             pass
            
        name, score, type_code, crop_processed = self.ocr.process_name_recognition(
            img_bgr, 
            rect, 
            self.matcher,
            ocr_mode=getattr(self.config, 'ocr_mode', 'offline'),
            preprocess_params=ATTENDANCE_PREPROCESS,
            online_crop_no_otsu=True,
            retry_with_shifts=True,
            item_id=f"b{block_idx}_r{row_idx}_c{col_idx}"
        )

        # 5. Сохранение отладочного изображения сопоставления
        if debug_dir and crop_processed is not None:
            safe_name = name if name else "UNKNOWN"
            # Очистка имени для использования в качестве названия файла
            safe_name = "".join(c for c in safe_name if c.isalnum() or c in (' ', '-', '_')).strip()
            debug_img_name = f"b{block_idx}_r{row_idx}_c{col_idx}_{safe_name}.jpg"
            debug_img_path = os.path.join(debug_dir, debug_img_name)
            try:
                Image.fromarray(crop_processed).save(debug_img_path)
            except Exception as e:
                self.logger.warning(f"Не удалось сохранить отладочное изображение {debug_img_path}: {e}")
                
        return name, score, type_code, x, curr_y

    def process_image(self, image_path, stop_event=None):
        """
        Обрабатывает одно изображение, используя многопоточность для отдельных ячеек.
        Возвращает список найденных имен.
        """
        if stop_event and stop_event.is_set():
            return []

        img = Image.open(image_path)
        draw = ImageDraw.Draw(img)
        # Пытаемся загрузить шрифт, иначе используем стандартный
        try:
            font = ImageFont.truetype("arial.ttf", self.grid_params['font_size'])
        except:
            font = ImageFont.load_default()

        found_names = []
        filename = os.path.basename(image_path)
        file_base_name = os.path.splitext(filename)[0]
        
        # Подготовка директории отладки, если требуется
        debug_dir = None
        if self.config.debug_screens:
            debug_dir = os.path.join(os.path.dirname(image_path), file_base_name)
            if os.path.exists(debug_dir):
                shutil.rmtree(debug_dir)
            os.makedirs(debug_dir, exist_ok=True)
            self.history.add_created(debug_dir)
            
        img_np = np.array(img)
        # Конвертация RGB в BGR для cv2
        img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        
        # У нас есть 2 блока: верхний и нижний (со смещением)
        shifts = [0, self.grid_params['shift_y']]
        
        tasks_args = []
        
        for block_idx, shift in enumerate(shifts):
            for row_idx, y in enumerate(self.grid_params['rows_y']):
                curr_y = y + shift
                for col_idx, x in enumerate(self.grid_params['cols_x']):
                    w = self.grid_params['name_w']
                    h = self.grid_params['name_h']
                    
                    # Проверка границ
                    if curr_y + h > img_np.shape[0] or x + w > img_np.shape[1]:
                         continue
                         
                    # Передаем полное изображение и координаты
                    # Примечание: img_bgr здесь не копируется, но OCR рассматривает его как доступный только для чтения
                    tasks_args.append((img_bgr, block_idx, row_idx, col_idx, x, curr_y, w, h, debug_dir))

        # Выполнение задач параллельно
        # Используем настроенное количество потоков или 8 по умолчанию
        num_threads = 8
        results = []
        
        if tasks_args:
             with concurrent.futures.ThreadPoolExecutor(max_workers=int(num_threads)) as executor:
                futures = [executor.submit(self._process_single_cell, *args) for args in tasks_args]
                for future in concurrent.futures.as_completed(futures):
                    if stop_event and stop_event.is_set():
                        # Отменяем ожидающие задачи и выходим
                        for f in futures:
                            f.cancel()
                        return []  # Возвращаем пустой список при остановке
                    try:
                        res = future.result()
                        results.append(res)
                    except Exception as e:
                        self.logger.error(f"Ошибка при обработке ячейки: {e}")
        
        # Обработка результатов
        for name, score, type_code, x, curr_y in results:
            if not name:
                continue

            if type_code != 3: # Не ошибка
                found_names.append(name)
            
            # Рисование на изображении
            text_x = x + 2
            text_y = curr_y + self.grid_params['shift_text_y']
            
            color = (255, 255, 255) # Белый по умолчанию
            if type_code == 1: # Нечеткое совпадение
                # Желтоватый цвет в зависимости от счета
                color = (255, 255, 160 - int((100 - score) * 4))
            elif type_code == 2: # Новый
                color = (7, 12, 180) # Синий
            elif type_code == 3: # Ошибка
                color = (187, 20, 20) # Красный
                name = "?????"
            elif type_code == 4: # Заменено
                color = (0, 0, 0) # Черный

            draw.text((text_x, text_y), name, font=font, fill=color)

        # Проверяем остановку перед сохранением
        if stop_event and stop_event.is_set():
            return found_names  # Возвращаем что успели собрать, но не сохраняем

        # Сохранение аннотированного изображения
        # Логика: сохранение в подпапку с датой
        try:
            date_folder = datetime.fromtimestamp(os.path.getmtime(image_path)).strftime("%Y-%m-%d")
            output_dir = os.path.join(os.path.dirname(image_path), date_folder)
            os.makedirs(output_dir, exist_ok=True)
            self.history.add_created(output_dir)
            
            base_name = os.path.basename(image_path)
            name_part, ext = os.path.splitext(base_name)
            output_path = os.path.join(output_dir, f"{name_part}_res{ext}")
            
            img.save(output_path)
            self.history.add_created(output_path)
            
            # Перемещение оригинала
            dest_path = os.path.join(output_dir, base_name)
            if os.path.exists(dest_path):
                os.remove(dest_path)
            shutil.move(image_path, dest_path)
            self.history.add_move(image_path, dest_path)
            
            # Перемещение папки отладки
            if debug_dir and os.path.exists(debug_dir):
                debug_dest_path = os.path.join(output_dir, os.path.basename(debug_dir))
                if os.path.exists(debug_dest_path):
                    shutil.rmtree(debug_dest_path)
                shutil.move(debug_dir, debug_dest_path)
                self.history.add_move(debug_dir, debug_dest_path)

            # Не открываем файл, если остановлено
            if self.config.show_afterscreen and not (stop_event and stop_event.is_set()):
                 os.startfile(output_path)

        except Exception as e:
             self.logger.error(f"Ошибка при сохранении результатов: {e}")

        return found_names
