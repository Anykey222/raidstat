import sys
import os

# Добавляем родительскую директорию в путь поиска модулей, чтобы Python видел пакет raidstat_py
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from raidstat_py.gui.app import App

def ensure_replacements_file():
    """Создает пустой файл Замены.txt рядом с исполняемым файлом, если он не существует."""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        # Для запуска как скрипт - корень проекта (на уровень выше raidstat_py)
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
    filename = os.path.join(base_path, "Замены.txt")
    if not os.path.exists(filename):
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                pass
        except Exception as e:
            print(f"Не удалось создать {filename}: {e}")

if __name__ == "__main__":
    ensure_replacements_file()
    app = App()
    app.mainloop()
