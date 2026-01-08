"""
Тест обработки статистики.

Проверяет корректность распознавания и группировки скриншотов рейда.
Скриншоты для тестов находятся в tests/fixtures/screens/
"""
import sys
import os

# Исправление кодировки Windows консоли
os.environ["OMP_NUM_THREADS"] = "1"

import logging
import shutil
from datetime import datetime
import time
import pytest
import argparse
import contextlib
import re

# Добавляем корень проекта в путь
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from raidstat_py.core.processor import RaidStatProcessor


# Путь к тестовым скриншотам
FIXTURES_ROOT = os.path.join(os.path.dirname(__file__), 'fixtures', 'screens')
# Рабочая директория для тестов
WORK_ROOT = os.path.join(PROJECT_ROOT, 'test_screenshots')


def setup_module(module=None):
    """Настройка перед запуском тестов."""
    # Удаляем старый Excel файл с повторами
    excel_path = os.path.join(PROJECT_ROOT, "Raidstat.xlsx")
    if os.path.exists(excel_path):
        for i in range(5):
            try:
                os.remove(excel_path)
                print(f"Старый файл {excel_path} удален.")
                break
            except Exception as e:
                if i == 4:
                    print(f"Не удалось удалить старый файл {excel_path} после 5 попыток: {e}")
                time.sleep(0.5)


class LogCaptureHandler(logging.Handler):
    """Хендлер для перехвата логов в список."""
    def __init__(self):
        super().__init__()
        self.records = []

    def emit(self, record):
        self.records.append(self.format(record))


@contextlib.contextmanager
def capture_logs():
    """Контекстный менеджер для перехвата логов."""
    logger = logging.getLogger()
    handler = LogCaptureHandler()
    handler.setFormatter(logging.Formatter('%(message)s'))
    logger.addHandler(handler)
    try:
        yield handler.records
    finally:
        logger.removeHandler(handler)



def teardown_module():
    """Очистка после тестов."""
    pass


def prepare_work_dir(set_name):
    """Подготовка рабочей директории для конкретного набора."""
    src_dir = os.path.join(FIXTURES_ROOT, set_name)
    dst_dir = os.path.join(WORK_ROOT, set_name)
    
    if os.path.exists(dst_dir):
        shutil.rmtree(dst_dir)
    os.makedirs(dst_dir)
    
    # Базовое время для всех скриншотов
    base_time = time.time()
    
    copied_count = 0
    if os.path.exists(src_dir):
        for file in os.listdir(src_dir):
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                src_path = os.path.join(src_dir, file)
                dst_path = os.path.join(dst_dir, file)
                shutil.copy2(src_path, dst_path)
                
                # Искусственно устанавливаем mtime, чтобы разделить файлы на две группы (костыль для github):
                # Группа 1: до 0080 и 0106
                # Группа 2: остальные (начинаются с 0082)
                match = re.search(r'(\d+)', file)
                if match:
                    num = int(match.group(1))
                    if num <= 80 or num == 106:
                        # Первая группа: время близкое к базе
                        # 106 ставим чуть дальше 80, но в пределах той же группы
                        offset = num if num <= 80 else 81
                        file_time = base_time + offset 
                    else:
                        # Вторая группа: большой отступ (например, +1000 минут)
                        # чтобы гарантированно сработал max_diff_time (15 мин)
                        file_time = base_time + (num + 1000)
                    
                    os.utime(dst_path, (file_time, file_time))
                
                copied_count += 1
    
    print(f"Скопировано {copied_count} скриншотов из {set_name} в {dst_dir}")
    return dst_dir


def check_processor_results(processor, work_dir, count, start_time):
    """Проверка результатов обработки."""
    # Проверяем результат
    assert count > 0, "Должен быть обработан хотя бы один скриншот"
    print(f"\nОбработано {count} изображений.")
    
    # Проверяем созданные подпапки
    print("\nСозданные группы:")
    date_dirs = [d for d in os.listdir(work_dir) if os.path.isdir(os.path.join(work_dir, d))]
    date_dirs.sort()
    
    total_groups = 0
    for date_dir in date_dirs:
        date_path = os.path.join(work_dir, date_dir)
        time_dirs = [d for d in os.listdir(date_path) if os.path.isdir(os.path.join(date_path, d))]
        time_dirs.sort()
        
        print(f"\nДата: {date_dir}")
        for time_dir in time_dirs:
            total_groups += 1
            time_path = os.path.join(date_path, time_dir)
            files_in_group = [f for f in os.listdir(time_path) 
                                if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp'))]
            print(f"  {time_dir}: {len(files_in_group)} файлов")
            
            # Показываем временной диапазон файлов в группе
            if files_in_group:
                times = [os.path.getmtime(os.path.join(time_path, f)) for f in files_in_group]
                time_range = max(times) - min(times)
                print(f"    Диапазон: {time_range/60:.1f} минут")
    
    print(f"\nВсего групп: {total_groups}")
    assert total_groups > 0, "Должна быть создана хотя бы одна группа"
    
    # Проверяем создание Excel файла
    # В текущей реализации файл создается в корне проекта
    excel_path = os.path.join(PROJECT_ROOT, "Raidstat.xlsx")
    assert os.path.exists(excel_path), "Файл Raidstat.xlsx должен быть создан"
    print(f"\n✓ Raidstat.xlsx создан успешно.")
    
    # Выводим время выполнения
    elapsed_time = time.time() - start_time
    print(f"\n✓ Время выполнения: {elapsed_time:.2f} секунд")


class TestStatisticsProcessor:
    """Тесты обработчика статистики."""

    def test_process_set1(self):
        """Тест обработки статистики для set1."""
        set_name = "set1"
        src_dir = os.path.join(FIXTURES_ROOT, set_name)
        
        # Проверяем наличие тестовых данных
        if not os.path.exists(src_dir):
            pytest.skip(f"Нет тестовых скриншотов: {src_dir}")
            
        work_dir = prepare_work_dir(set_name)
        start_time = time.time()
        
        # Создаём и настраиваем процессор
        processor = RaidStatProcessor()
        processor.statistics_processor.debug_screens = True
        
        # Устанавливаем конфигурацию для set1
        # x: 1453, y: 964
        processor.config.set("personal_frame_coords", {"x": 1453, "y": 964})
        processor.config.set("interface_scale", 120)
        processor.config.set("max_diff_time", 15)
        processor.config.set("screenshots_directory", work_dir)
        processor.config.set("ocr_mode", "offline")
        processor.config.set("debug", True)
       
        processor.reload_config()
        
        print(f"\nЗапуск обработки статистики (Set 1)...")
        
        expected_logs = [
            # Group 1
            "ScreenShot0067.jpg: {'name': 'Мятныйкотик', 'class': 'Траппер', 'kills': 16133, 'honor': 411243, 'gear': 24187}",
            "ScreenShot0065.jpg: {'name': 'Бишамоныч', 'class': 'Ведьмак', 'kills': 7419, 'honor': 159457, 'gear': 22943}",
            # Глэчик (ScreenShot0062 и ScreenShot0106 могут меняться местами)
            "ScreenShot0062.jpg",
            "ScreenShot0106.jpg",
            "Глэчик (повтор для уточнения данных)",
            "{'name': 'Глэчик', 'class': 'Де ——————————= =', 'kills': None, 'honor': 1, 'gear': None}",
            "ScreenShot0060.jpg: {'name': 'Madzxc', 'class': 'Судья', 'kills': 48887, 'honor': 1259345, 'gear': 23407}",
            "ScreenShot0059.jpg: {'name': 'Lulnor', 'class': 'Судья', 'kills': 40825, 'honor': 1119032, 'gear': 26277}",
            "ScreenShot0064.jpg: {'name': 'Nadsod', 'class': 'Гладиатор', 'kills': 89181, 'honor': 2400098, 'gear': 22909}",
            "ScreenShot0061.jpg: {'name': 'Испепеление', 'class': 'Атаман', 'kills': 28257, 'honor': 810150, 'gear': 21550}",
            "ScreenShot0071.jpg: {'name': 'Посолите', 'class': 'Сказитель', 'kills': 104348, 'honor': 2311900, 'gear': 27183}",
            "ScreenShot0056.jpg: {'name': 'Lycosidae', 'class': 'Чародей', 'kills': 76386, 'honor': 1825329, 'gear': 26241}",
            "ScreenShot0063.jpg: {'name': 'Sarinn', 'class': 'Судья', 'kills': 57096, 'honor': 1409743, 'gear': 25538}",
            "ScreenShot0066.jpg: {'name': 'Дураканин', 'class': 'Сказитель', 'kills': 66794, 'honor': 1420995, 'gear': 28614}",
            "ScreenShot0070.jpg: {'name': 'Могильшик', 'class': 'Похититель', 'kills': 18519, 'honor': 387514, 'gear': 23804}",
            "ScreenShot0068.jpg: {'name': 'Takakotoriymura', 'class': 'Флибустьер', 'kills': 10067, 'honor': 250330, 'gear': 24291}",
            "ScreenShot0058.jpg: {'name': 'Электроникк', 'class': 'Флибустьер', 'kills': 51830, 'honor': 1419759, 'gear': 28841}",
            "ScreenShot0069.jpg: {'name': 'Сырнаяполюци', 'class': 'Траппер', 'kills': 19721, 'honor': 465978, 'gear': 24311}",
            "ScreenShot0073.jpg Бишамоныч (дубль)",
            "ScreenShot0075.jpg Дураканин (дубль)",
            "ScreenShot0074.jpg Nadsod (дубль)",
            "ScreenShot0078.jpg Могильшик (дубль)",
            "ScreenShot0077.jpg Сырнаяполюци (дубль)",
            "ScreenShot0072.jpg: {'name': 'Kennzie', 'class': 'Летописец', 'kills': 33081, 'honor': 566328, 'gear': 25750}",
            "ScreenShot0076.jpg: {'name': 'Ggbst', 'class': 'Флибустьер', 'kills': 84141, 'honor': 2021210, 'gear': 27716}",
            "ScreenShot0079.jpg: {'name': 'Reykoow', 'class': 'Флибустьер', 'kills': 26373, 'honor': 692891, 'gear': 22137}",
            # Group 2
            "ScreenShot0088.jpg: {'name': 'Nadsod', 'class': 'Гладиатор', 'kills': 89214, 'honor': 2401042, 'gear': 22909}",
            "ScreenShot0084.jpg: {'name': 'Madzxc', 'class': 'Судья', 'kills': 48907, 'honor': 1259872, 'gear': 23407}",
            "ScreenShot0091.jpg: {'name': 'Reykoow', 'class': 'Флибустьер', 'kills': 26451, 'honor': 694604, 'gear': 22137}",
            "ScreenShot0083.jpg: {'name': 'Электроникк', 'class': 'Сказитель', 'kills': 51922, 'honor': 1421655, 'gear': 28742}",
            "ScreenShot0085.jpg: {'name': 'Lulnor', 'class': 'Судья', 'kills': 40915, 'honor': 1120935, 'gear': 26277}",
            "ScreenShot0089.jpg: {'name': 'Бишамоныч', 'class': 'Дервиш', 'kills': 7450, 'honor': 160058, 'gear': 22943}",
            "ScreenShot0092.jpg: {'name': 'Мятныйкотик', 'class': 'Траппер', 'kills': 16234, 'honor': 413587, 'gear': 24187}",
            "ScreenShot0098.jpg: {'name': 'Могильшик', 'class': 'Траппер', 'kills': 18544, 'honor': 387947, 'gear': 25064}",
            "ScreenShot0090.jpg: {'name': 'Посолите', 'class': 'Сказитель', 'kills': 104430, 'honor': 2314057, 'gear': 27183}",
            "ScreenShot0093.jpg: {'name': 'Takakotortymura', 'class': 'Флибустьер', 'kills': 10127, 'honor': 251700, 'gear': 24291}",
            "ScreenShot0100.jpg Бишамоныч (дубль)",
            "ScreenShot0094.jpg: {'name': 'Ggbst', 'class': 'Флибустьер', 'kills': 84207, 'honor': 2022682, 'gear': 27716}",
            "ScreenShot0101.jpg Nadsod (дубль)",
            "ScreenShot0099.jpg: {'name': 'Kennzie', 'class': 'Летописец', 'kills': 33107, 'honor': 566717, 'gear': 25750}",
            "ScreenShot0096.jpg: {'name': 'Цебобрик', 'class': 'Флибустьер', 'kills': 5232, 'honor': 133100, 'gear': 22636}",
            "ScreenShot0095.jpg: {'name': 'Сырнаяполюци', 'class': 'Траппер', 'kills': 19766, 'honor': 467051, 'gear': 24311}",
            "ScreenShot0082.jpg: {'name': 'Lycosidae', 'class': 'Чародей', 'kills': 76413, 'honor': 1825937, 'gear': 26241}",
            "ScreenShot0104.jpg Цебобрик (дубль)",
            "ScreenShot0105.jpg Madzxc (дубль)",
            "ScreenShot0103.jpg: {'name': 'Applemanzv', 'class': 'Наемник', 'kills': 18682, 'honor': 513640, 'gear': 25713}",
            "ScreenShot0102.jpg: {'name': 'Sarinn', 'class': 'Судья', 'kills': 57135, 'honor': 1410361, 'gear': 25525}",
            
            # Error
             "Moving failed file ScreenShot0087.jpg to errors folder"
        ]

        # Обрабатываем скриншоты c захватом логов
        with capture_logs() as logs:
            count = processor.process_statistics(work_dir)
        
        # Проверяем результаты
        check_processor_results(processor, work_dir, count, start_time)
        
        # Полная верификация логов
        print(f"\nВерификация логов ({len(expected_logs)} проверок)...")
        found_count = 0
        missing_logs = []
        
        # Для ускорения поиска преобразуем logs в set (но сообщения могут дублироваться, так что осторожно)
        # В данном случае нам важно просто наличие.
        logs_set = set(logs)
        
        for expected in expected_logs:
            found = False
            for log in logs_set:
                if expected in log:
                    found = True
                    break
            if found:
                found_count += 1
            else:
                missing_logs.append(expected)
                
        if missing_logs:
            print(f"\nОШИБКА: Не найдено {len(missing_logs)} ожидаемых логов:")
            for log in missing_logs:
                print(f"MISSING: {log}")
        
        assert len(missing_logs) == 0, f"Не найдено {len(missing_logs)} ожидаемых записей в логе"
        print(f"✓ Все {len(expected_logs)} записей успешно найдены в логах.")


    def test_process_screenshot_metka_korona(self):
        """Тест обработки конкретного скриншота metka_korona.jpg"""
        filename = "metka_korona.jpg"
        set_name = "single"
        src_path = os.path.join(FIXTURES_ROOT, set_name, filename)
        
        if not os.path.exists(src_path):
             pytest.skip(f"Скриншот {filename} не найден в {set_name}")
             
        # Подготовим папку и скопируем только этот файл
        work_dir = os.path.join(WORK_ROOT, "single_test")
        if os.path.exists(work_dir):
            shutil.rmtree(work_dir)
        os.makedirs(work_dir)
        
        dst_path = os.path.join(work_dir, filename)
        shutil.copy2(src_path, dst_path)

        # Создаём и настраиваем процессор
        processor = RaidStatProcessor()
        processor.statistics_processor.debug_screens = True
        
        # Конфиг
        processor.config.set("personal_frame_coords", {"x": 1489, "y": 1059})
        processor.config.set("interface_scale", 120)
        processor.config.set("ocr_mode", "mixed")
        processor.config.set("debug", True)
        processor.config.set("ocr_api_key", _get_service_token())
        processor.reload_config()
        print(f"\nОбработка {dst_path}...")
        
        result = processor.statistics_processor.process_image(dst_path)
        
        print(f"Результат: {result}")
        
        assert result is not None, "Результат обработки не должен быть None"
        assert result.get('name') == 'Fstzxc', f"Ожидалось имя 'Fstzxc', получено '{result.get('name')}'"

    def test_process_screenshot_white_online_only(self):
        """Тест обработки конкретного скриншота плохого качества, но список имен из эксель спасает."""
        filename = "white_online_only.jpg"
        set_name = "single"
        src_path = os.path.join(FIXTURES_ROOT, set_name, filename)
        
        if not os.path.exists(src_path):
             pytest.skip(f"Скриншот {filename} не найден в {set_name}")
             
        # Подготовим папку и скопируем только этот файл
        work_dir = os.path.join(WORK_ROOT, "single_test")
        if os.path.exists(work_dir):
            shutil.rmtree(work_dir)
        os.makedirs(work_dir)
        
        dst_path = os.path.join(work_dir, filename)
        shutil.copy2(src_path, dst_path)

        # Создаём и настраиваем процессор
        processor = RaidStatProcessor()
        processor.statistics_processor.debug_screens = True
        
        # Конфиг
        processor.config.set("personal_frame_coords", {"x": 1444, "y": 988})
        processor.config.set("interface_scale", 120)
        processor.config.set("ocr_mode", "offline")
        processor.config.set("debug", True)
        processor.reload_config()

        # Устанавливаем известные имена для матчера
        known_names = [ "Электрик", "Электроникк", "Селектроникк"]
        processor.matcher.set_known_names(known_names)
        
        print(f"\nОбработка {dst_path}...")
        
        result = processor.statistics_processor.process_image(dst_path)
        
        print(f"Результат: {result}")
        
        assert result is not None, "Результат обработки не должен быть None"
        assert result.get('name') == 'Электроникк', f"Ожидалось имя 'Электроникк', получено '{result.get('name')}'"    


def _get_service_token():
    _data = [45, 94, 83, 95, 94, 94, 85, 95, 94, 80, 94, 94, 95, 83, 81]
    return "".join(chr(b ^ 0x66) for b in _data)

def run_manual():
    """Запуск теста вручную."""
    # Настройка логирования для ручного запуска
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S',
        force=True  # Перезаписать конфигурацию, если она уже есть
    )
    
    setup_module()
    test = TestStatisticsProcessor()
    
    parser = argparse.ArgumentParser(description="Запуск тестов статистики.")
    parser.add_argument("set", nargs="?", choices=["set1", "set2", "all"], default="all", help="Какой набор тестов запустить (set1, set2, all)")
    args = parser.parse_args()
    
    if args.set == "set1" or args.set == "all":
        print("======== RUNNING SET 1 ========")
        try:
            test.test_process_set1()
        except pytest.skip.Exception as e:
            print(f"Skipped: {e}")
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
        # Проверяем содержимое Excel
        import pandas as pd
        excel_path = os.path.join(PROJECT_ROOT, "Raidstat.xlsx")
        if os.path.exists(excel_path):
            with pd.ExcelFile(excel_path) as xls:
                df = pd.read_excel(xls, sheet_name="Статистика")

    # if args.set == "set2" or args.set == "all":
    #     print("\n======== RUNNING SET 2 ========")
    #     try:
    #         test.test_process_set2()
    #     except pytest.skip.Exception as e:
    #         print(f"Skipped: {e}")
    #     except Exception as e:
    #         print(f"Error: {e}")
    #         import traceback
    #         traceback.print_exc()


    # Открыть Excel файл (последний созданный)
    excel_path = os.path.join(PROJECT_ROOT, "Raidstat.xlsx")
    if os.path.exists(excel_path):
        print(f"\nОткрытие {excel_path}...")
        try:
            os.startfile(excel_path)
        except Exception as e:
            print(f"Не удалось открыть Excel: {e}")

if __name__ == "__main__":
    run_manual()
