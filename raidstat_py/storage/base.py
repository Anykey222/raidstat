from abc import ABC, abstractmethod

class StorageInterface(ABC):
    @abstractmethod
    def get_roster(self):
        """Возвращает список известных имен."""
        pass

    @abstractmethod
    def save_attendance(self, attendance_data, date_str):
        """Сохраняет данные о посещаемости."""
        pass

    @abstractmethod
    def save_statistics(self, stats_data, date_str):
        """Сохраняет данные статистики."""
        pass
