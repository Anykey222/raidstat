import json
import os
import shutil
import logging
import threading

class HistoryManager:
    def __init__(self, mode):
        self.mode = mode
        self.filename = f"history_{mode}.json"
        self.moves = []
        self.created = []
        self.logger = logging.getLogger(__name__)
        self.lock = threading.Lock()

    def clear(self):
        with self.lock:
            self.moves = []
            self.created = []
        self.save()

    def add_move(self, src, dest):
        with self.lock:
            self.moves.append({"src": src, "dest": dest})

    def add_created(self, path):
        with self.lock:
            self.created.append(path)

    def save(self):
        data = {
            "moves": self.moves,
            "created": self.created
        }
        try:
            with open(self.filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.error(f"Не удалось сохранить историю: {e}")

    def load(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.moves = data.get("moves", [])
                    self.created = data.get("created", [])
            except Exception as e:
                self.logger.error(f"Не удалось загрузить историю: {e}")
        else:
            self.moves = []
            self.created = []

    def revert(self):
        self.load()
        if not self.moves and not self.created:
            raise Exception("Нет действий для отмены (история пуста или отсутствует).")

        self.logger.info(f"Откат {len(self.moves)} перемещений и {len(self.created)} созданных файлов...")

        # Откат перемещений
        for move in reversed(self.moves):
            src = move['src']
            dest = move['dest']
            
            if os.path.exists(dest):
                try:
                    os.makedirs(os.path.dirname(src), exist_ok=True)
                    if os.path.exists(src):
                        os.remove(src)
                    shutil.move(dest, src)
                    self.logger.info(f"Восстановлено: {os.path.basename(dest)} -> {os.path.dirname(src)}")
                except Exception as e:
                    self.logger.error(f"Не удалось восстановить {dest} -> {src}: {e}")
            else:
                self.logger.warning(f"Файл не найден для отката: {dest}")

        # Удаление созданных файлов/директорий
        for path in self.created:
            if os.path.exists(path):
                try:
                    if os.path.isdir(path):
                        shutil.rmtree(path)
                    else:
                        os.remove(path)
                    self.logger.info(f"Удалено: {path}")
                except Exception as e:
                    self.logger.error(f"Не удалось удалить {path}: {e}")

        self.clear()
        self.logger.info("Откат завершен.")
