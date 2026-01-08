import json
import os
import logging

class Config:
    DEFAULT_CONFIG = {
        "interface_scale": 120,
        "raid_frame_coords": {"x": 1394, "y": 1016},
        "personal_frame_coords": {"x": 1453, "y": 964},
        "max_diff_time": 15,
        "screenshots_directory": "",  # Последняя выбранная директория со скриншотами
        "show_afterscreen": False,
        "debug_screens": False,
        "recursive_scan": False,
        "ocr_mode": "offline",
        "ocr_api_key": "",
        "debug": False
    }

    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self.data = self.DEFAULT_CONFIG.copy()
        self.load()

    def load(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    self.data.update(loaded)
            except Exception as e:
                logging.error(f"Ошибка загрузки конфигурации: {e}")
        else:
            self.save()

    def save(self):
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Ошибка сохранения конфигурации: {e}")

    def get(self, key):
        return self.data.get(key, self.DEFAULT_CONFIG.get(key))

    def set(self, key, value):
        self.data[key] = value
        self.save()

    # Типизированные геттеры для удобства
    @property
    def interface_scale(self): return int(self.data.get("interface_scale", 120))
    
    @property
    def raid_frame_coords(self): return self.data.get("raid_frame_coords")
    
    @property
    def personal_frame_coords(self): return self.data.get("personal_frame_coords")

    @property
    def max_diff_time(self): return int(self.data.get("max_diff_time", 15))

    @property
    def show_afterscreen(self): return bool(self.data.get("show_afterscreen", False))

    @property
    def debug_screens(self): return bool(self.data.get("debug_screens", False))

    @property
    def recursive_scan(self): return bool(self.data.get("recursive_scan", False))

    @property
    def ocr_mode(self): return self.data.get("ocr_mode", "offline")

    @property
    def ocr_api_key(self): return self.data.get("ocr_api_key", "")

    @property
    def debug(self): return bool(self.data.get("debug", False))
