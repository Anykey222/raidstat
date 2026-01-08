import logging
import threading
from ..utils.config import Config
from .ocr import OCRHandler
from .matcher import Matcher
from ..storage.excel_impl import ExcelStorage
from .attendance import AttendanceProcessor
from .statistics import StatisticsProcessor

class RaidStatProcessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # Инициализация компонентов
        self.config = Config()
        self.ocr = OCRHandler(config=self.config)
        self.storage = ExcelStorage()
        
        # Загрузка ростера для матчера
        roster = self.storage.get_roster()
        self.matcher = Matcher(known_names=roster)
        
        # Определяем путь к Замены.txt
        import sys
        import os
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            # core -> raidstat_py -> raidstat (root)
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        
        self.matcher.load_replacements(os.path.join(base_path, "Замены.txt"))
        
        # Процессоры
        self.attendance_processor = AttendanceProcessor(self.config, self.ocr, self.matcher, self.storage)
        self.statistics_processor = StatisticsProcessor(self.config, self.ocr, self.matcher, self.storage, debug_screens=self.config.get("debug_screens"))
        
        # Контроль
        self.stop_event = threading.Event()

    def revert_attendance(self):
        self.attendance_processor.revert_history()

    def revert_statistics(self):
        self.statistics_processor.revert_history()

    def stop_processing(self):
        self.logger.info("Остановка обработки...")
        self.stop_event.set()

    def process_attendance(self, folder_path, recursive=False):
        self.stop_event.clear()
        # Перезагружаем ростер, чтобы использовать актуальные имена из Excel
        self.matcher.set_known_names(self.storage.get_roster(source="attendance"))
        
        self.logger.info(f"Начало обработки посещаемости в {folder_path}")
        return self.attendance_processor.process_folder(folder_path, recursive, self.stop_event)

    def process_statistics(self, folder_path, recursive=False):
        self.stop_event.clear()
        # Перезагружаем ростер здесь тоже
        self.matcher.set_known_names(self.storage.get_roster(source="statistics"))
        
        self.logger.info(f"Начало сбора статистики в {folder_path}")
        return self.statistics_processor.process_folder(folder_path, recursive, self.stop_event)

    def reload_config(self):
        self.config.load()
        # Обновляем процессоры при необходимости (они ссылаются на объект конфига, так что должно быть норм)
        # Но параметры сетки могут нуждаться в пересчете
        self.attendance_processor.grid_params = self.attendance_processor._get_grid_params(self.config.interface_scale)
        self.statistics_processor.offsets = self.statistics_processor._get_offsets(self.config.interface_scale)
        self.statistics_processor.debug_screens = self.config.get("debug_screens")
