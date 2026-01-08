import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import logging
import threading
import sys
import os
from ..core.processor import RaidStatProcessor
from .cropper import CropWindow
import webbrowser

# --- COLORS & THEME ---
class Theme:
    BG_MAIN = "#0f172a"      # Slate 900
    BG_CARD = "#1e293b"      # Slate 800
    BG_SIDEBAR = "#1e293b"   # Slate 800
    
    TEXT_PRIMARY = "#f8fafc" # Slate 50
    TEXT_SECONDARY = "#94a3b8" # Slate 400
    
    ACCENT_BLUE = "#3b82f6"  # Blue 500
    ACCENT_GREEN = "#10b981" # Emerald 500
    ACCENT_RED = "#ef4444"   # Red 500
    
    BTN_HOVER_BLUE = "#2563eb"
    BTN_HOVER_GREEN = "#059669"
    BTN_HOVER_RED = "#dc2626"
    
    CONSOLE_BG = "#020617"   # Slate 950
    CONSOLE_TEXT = "#22c55e" # Green terminal text

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue") # Basis, we override mostly

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            try:
                self.text_widget.configure(state='normal')
                self.text_widget.insert('end', msg + '\n')
                self.text_widget.configure(state='disabled')
                self.text_widget.see('end')
            except Exception:
                pass
        self.text_widget.after(0, append)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("RaidStat")
        self.geometry("1150x800")
        self.configure(fg_color=Theme.BG_MAIN) # Main window background
        
        # Инициализация процессора
        self.processor = RaidStatProcessor()

        # Установка иконки
        self._set_icon()

        # Настройка сетки окна
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Боковая панель (Sidebar) ---
        self.sidebar_frame = ctk.CTkFrame(self, width=240, corner_radius=0, fg_color=Theme.BG_SIDEBAR)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1) # Spacer

        # Логотип
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="RaidStat", 
            font=ctk.CTkFont(family="Roboto Medium", size=26, weight="bold"),
            text_color=Theme.TEXT_PRIMARY
        )
        self.logo_label.grid(row=0, column=0, padx=25, pady=(35, 20), sticky="w")

        # Кнопки навигации
        self.nav_buttons = {}
        self.create_nav_btn("📊  Посещаемость", self.show_attendance_tab, "attendance", 1)
        self.create_nav_btn("📈  Статистика", self.show_statistics_tab, "statistics", 2)
        self.create_nav_btn("⚙️  Настройки", self.show_settings_tab, "settings", 3)
        self.create_nav_btn("ℹ️  О проекте", self.show_about_tab, "about", 4)

        # Версия
        self.version_label = ctk.CTkLabel(self.sidebar_frame, text="v1.0", text_color=Theme.TEXT_SECONDARY)
        self.version_label.grid(row=6, column=0, padx=20, pady=20)

        # --- Основная область контента ---
        self.main_panel = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_panel.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        self.main_panel.grid_rowconfigure(0, weight=0) 
        self.main_panel.grid_rowconfigure(1, weight=1)
        self.main_panel.grid_columnconfigure(0, weight=1)

        # Контейнер для вкладок
        self.content_container = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        self.content_container.grid(row=0, column=0, sticky="nsew", padx=30, pady=30)
        self.content_container.grid_columnconfigure(0, weight=1)
        self.content_container.grid_rowconfigure(0, weight=1)

        # Фреймы разделов (инициализируем пустые, наполняем позже)
        self.frame_attendance = None
        self.frame_statistics = None
        self.frame_settings = None
        self.frame_about = None

        # Инициализация всех вкладок
        self.setup_tabs()

        # --- Панель логов (внизу) ---
        self.console_frame = ctk.CTkFrame(self.main_panel, height=200, fg_color=Theme.BG_CARD, corner_radius=10)
        self.console_frame.grid(row=1, column=0, sticky="nsew", padx=30, pady=(0, 30))
        self.console_frame.grid_columnconfigure(0, weight=1)
        self.console_frame.grid_rowconfigure(1, weight=1)

        # Заголовок консоли
        self.console_header = ctk.CTkFrame(self.console_frame, fg_color="transparent", height=30)
        self.console_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(10,0))
        
        ctk.CTkLabel(
            self.console_header, 
            text="💻 Лог работы", 
            text_color=Theme.TEXT_SECONDARY,
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left")

        # Текстовое поле консоли
        self.console_text = ctk.CTkTextbox(
            self.console_frame, 
            height=150, 
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color=Theme.CONSOLE_BG,
            text_color=Theme.CONSOLE_TEXT,
            activate_scrollbars=True
        )
        self.console_text.grid(row=1, column=0, padx=15, pady=10, sticky="nsew")
        self.console_text.configure(state='disabled')

        # Бинды для консоли
        self.console_text.bind("<KeyPress>", self._handle_console_keypress)

        # Настройка логгирования
        self.setup_logging()

        # Открываем первую вкладку
        self.show_attendance_tab()

    def create_nav_btn(self, text, command, name, row):
        btn = ctk.CTkButton(
            self.sidebar_frame, 
            text=text, 
            command=command, 
            font=ctk.CTkFont(size=15, weight="normal"),
            fg_color="transparent",
            text_color=Theme.TEXT_SECONDARY,
            hover_color=Theme.BG_MAIN,
            anchor="w",
            height=45,
            corner_radius=8
        )
        btn.grid(row=row, column=0, padx=15, pady=5, sticky="ew")
        self.nav_buttons[name] = btn

    def setup_tabs(self):
        # Посещаемость
        self.frame_attendance = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.setup_attendance_view(self.frame_attendance)
        
        # Статистика
        self.frame_statistics = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.setup_statistics_view(self.frame_statistics)
        
        # Настройки
        self.frame_settings = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.setup_settings_view(self.frame_settings)
        
        # О проекте
        self.frame_about = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.setup_about_view(self.frame_about)

    def setup_logging(self):
        # Очищаем существующие хендлеры
        root_logger = logging.getLogger()
        root_logger.handlers = []
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # 1. GUI Handler
        text_handler = TextHandler(self.console_text)
        text_handler.setFormatter(formatter)
        root_logger.addHandler(text_handler)
        
        # 2. File Handler
        try:
            file_handler = logging.FileHandler("raidstat.log", encoding='utf-8')
            file_handler.setFormatter(formatter)
            root_logger.addHandler(file_handler)
        except Exception as e:
            print(f"Не удалось создать файл лога: {e}")

        # 3. Stream Handler
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)
        root_logger.addHandler(stream_handler)
        
        debug_mode = self.processor.config.debug
        root_logger.setLevel(logging.DEBUG if debug_mode else logging.INFO)

    def _set_icon(self):
        try:
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, "gui", "assets", "icon.ico")
            else:
                current_dir = os.path.dirname(os.path.abspath(__file__))
                icon_path = os.path.join(current_dir, "assets", "icon.ico")
            
            if os.path.exists(icon_path):
                self.after(200, lambda: self.iconbitmap(icon_path))
        except Exception as e:
            logging.warning(f"Не удалось установить иконку: {e}")

    def highlight_btn(self, btn_name):
        for name, btn in self.nav_buttons.items():
            if name == btn_name:
                btn.configure(fg_color=Theme.BG_MAIN, text_color=Theme.TEXT_PRIMARY, font=ctk.CTkFont(size=15, weight="bold"))
            else:
                btn.configure(fg_color="transparent", text_color=Theme.TEXT_SECONDARY, font=ctk.CTkFont(size=15, weight="normal"))

    def show_attendance_tab(self):
        self._show_frame(self.frame_attendance, "attendance")

    def show_statistics_tab(self):
        self._show_frame(self.frame_statistics, "statistics")

    def show_settings_tab(self):
        self._show_frame(self.frame_settings, "settings")

    def show_about_tab(self):
        self._show_frame(self.frame_about, "about")

    def _show_frame(self, frame, name):
        self.highlight_btn(name)
        # Скрываем все
        for f in [self.frame_attendance, self.frame_statistics, self.frame_settings, self.frame_about]:
            if f: f.grid_forget()
        
        # Настройка весов: в настройках приоритет контенту, в остальных - логу
        if name == "settings":
            self.main_panel.grid_rowconfigure(0, weight=1)
            self.main_panel.grid_rowconfigure(1, weight=0)
        else:
            self.main_panel.grid_rowconfigure(0, weight=0)
            self.main_panel.grid_rowconfigure(1, weight=1)

        # Показываем нужный
        if frame:
            frame.grid(row=0, column=0, sticky="nsew")

    # --- UI Components Helpers ---
    def create_card(self, parent, title=None):
        card = ctk.CTkFrame(parent, fg_color=Theme.BG_CARD, corner_radius=15)
        if title:
            lbl = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=18, weight="bold"), text_color=Theme.TEXT_PRIMARY)
            lbl.pack(padx=20, pady=(20, 10), anchor="w")
        return card

    def create_section_header(self, parent, text):
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="x", pady=(20, 10))
        lbl = ctk.CTkLabel(container, text=text, font=ctk.CTkFont(size=14, weight="bold"), text_color=Theme.ACCENT_BLUE)
        lbl.pack(side="left", padx=5)
        separator = ctk.CTkFrame(container, height=2, fg_color=Theme.BG_MAIN)
        separator.pack(side="left", fill="x", expand=True, padx=10)
        return container

    # --- Views ---
    
    def setup_attendance_view(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        
        # Заголовок
        ctk.CTkLabel(parent, text="Учет посещаемости", font=ctk.CTkFont(size=24, weight="bold"), text_color=Theme.TEXT_PRIMARY).grid(row=0, column=0, padx=0, pady=(0, 20), sticky="w")
        
        # Карточка управления
        card = self.create_card(parent)
        card.grid(row=1, column=0, sticky="ew")
        
        # Выбор папки
        ctk.CTkLabel(card, text="Папка со скриншотами:", text_color=Theme.TEXT_SECONDARY).pack(padx=20, pady=(20, 5), anchor="w")
        
        path_frame = ctk.CTkFrame(card, fg_color="transparent")
        path_frame.pack(padx=20, fill="x")
        
        self.att_folder_path = ctk.StringVar(value=self.processor.config.get("screenshots_directory") or "Папка не выбрана")
        entry = ctk.CTkEntry(
            path_frame, 
            textvariable=self.att_folder_path, 
            state="readonly", 
            placeholder_text="Путь к папке...",
            height=40,
            fg_color=Theme.BG_MAIN,
            border_width=1,
            border_color="#334155"
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_select = ctk.CTkButton(
            path_frame, 
            text="📂 Выбрать", 
            command=lambda: self.select_folder(self.att_folder_path), 
            width=100, 
            height=40,
            fg_color=Theme.ACCENT_BLUE,
            hover_color=Theme.BTN_HOVER_BLUE
        )
        btn_select.pack(side="left")

        # Опции
        self.att_recursive = ctk.BooleanVar(value=self.processor.config.recursive_scan)
        chk_recursive = ctk.CTkCheckBox(
            card, 
            text="🔍 Искать в подпапках", 
            variable=self.att_recursive, 
            command=self.save_attendance_options,
            text_color=Theme.TEXT_PRIMARY,
            hover_color=Theme.ACCENT_BLUE,
            fg_color=Theme.ACCENT_BLUE
        )
        chk_recursive.pack(padx=20, pady=20, anchor="w")

        # Кнопки действий
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=20, sticky="ew")
        
        self.btn_att_process = ctk.CTkButton(
            btn_frame, 
            text="🚀 Начать обработку", 
            command=self.run_attendance,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            fg_color=Theme.ACCENT_GREEN,
            hover_color=Theme.BTN_HOVER_GREEN
        )
        self.btn_att_process.pack(fill="x", pady=(0, 10))
        
        btn_revert = ctk.CTkButton(
            btn_frame, 
            text="↩️ Вернуть скриншоты", 
            command=self.revert_attendance_ui,
            fg_color="transparent",
            border_width=1,
            border_color=Theme.ACCENT_RED,
            text_color=Theme.ACCENT_RED,
            hover_color=Theme.BG_CARD,
            height=35
        )
        btn_revert.pack(fill="x")

    def setup_statistics_view(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(parent, text="Сбор статистики", font=ctk.CTkFont(size=24, weight="bold"), text_color=Theme.TEXT_PRIMARY).grid(row=0, column=0, padx=0, pady=(0, 20), sticky="w")
        
        card = self.create_card(parent)
        card.grid(row=1, column=0, sticky="ew")
        
        ctk.CTkLabel(card, text="Папка со скриншотами:", text_color=Theme.TEXT_SECONDARY).pack(padx=20, pady=(20, 5), anchor="w")
        
        path_frame = ctk.CTkFrame(card, fg_color="transparent")
        path_frame.pack(padx=20, fill="x", pady=(0, 20))
        
        self.stat_folder_path = ctk.StringVar(value=self.processor.config.get("screenshots_directory") or "Папка не выбрана")
        entry = ctk.CTkEntry(
            path_frame, 
            textvariable=self.stat_folder_path, 
            state="readonly",
            height=40,
            fg_color=Theme.BG_MAIN,
            border_color="#334155"
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_select = ctk.CTkButton(
            path_frame, 
            text="📂 Выбрать", 
            command=lambda: self.select_folder(self.stat_folder_path), 
            width=100, 
            height=40,
            fg_color=Theme.ACCENT_BLUE,
            hover_color=Theme.BTN_HOVER_BLUE
        )
        btn_select.pack(side="left")

        # Actions
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=2, column=0, pady=20, sticky="ew")

        self.btn_stat_process = ctk.CTkButton(
            btn_frame, 
            text="📊 Собрать статистику", 
            command=self.run_statistics,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            fg_color=Theme.ACCENT_GREEN,
            hover_color=Theme.BTN_HOVER_GREEN
        )
        self.btn_stat_process.pack(fill="x", pady=(0, 10))
        
        btn_revert = ctk.CTkButton(
            btn_frame, 
            text="↩️ Вернуть скриншоты", 
            command=self.revert_statistics_ui,
            fg_color="transparent",
            border_width=1,
            border_color=Theme.ACCENT_RED,
            text_color=Theme.ACCENT_RED,
            hover_color=Theme.BG_CARD,
            height=35
        )
        btn_revert.pack(fill="x")

    def setup_settings_view(self, parent):
        # Используем прокрутку для настроек, если их будет много
        scroll_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll_frame.pack(fill="both", expand=True)
        
        ctk.CTkLabel(scroll_frame, text="Настройки программы", font=ctk.CTkFont(size=24, weight="bold"), text_color=Theme.TEXT_PRIMARY).pack(anchor="w", pady=(0, 20))
        
        # --- Интерфейс ---
        self.create_section_header(scroll_frame, "ВИЗУАЛЬНЫЕ НАСТРОЙКИ")
        card_ui = self.create_card(scroll_frame)
        card_ui.pack(fill="x", pady=10)
        
        # Grid для UI
        grid_ui = ctk.CTkFrame(card_ui, fg_color="transparent")
        grid_ui.pack(padx=20, pady=20, fill="x")
        
        ctk.CTkLabel(grid_ui, text="Масштаб интерфейса:", text_color=Theme.TEXT_SECONDARY).grid(row=0, column=0, sticky="w", pady=10)
        self.var_scale = ctk.StringVar(value=str(self.processor.config.interface_scale))
        ctk.CTkComboBox(grid_ui, values=["100", "110", "120", "130"], variable=self.var_scale, command=self.save_settings, width=150).grid(row=0, column=1, sticky="w", padx=20)

        # --- Распознавание ---
        self.create_section_header(scroll_frame, "ОБЛАСТИ РАСПОЗНАВАНИЯ")
        card_ocr = self.create_card(scroll_frame)
        card_ocr.pack(fill="x", pady=10)
        
        grid_ocr = ctk.CTkFrame(card_ocr, fg_color="transparent")
        grid_ocr.pack(padx=20, pady=20, fill="x")
        
        # Рейд
        ctk.CTkLabel(grid_ocr, text="Координаты Рейда (X, Y):", text_color=Theme.TEXT_SECONDARY).grid(row=0, column=0, sticky="w", pady=10)
        self.lbl_raid_coords = ctk.CTkLabel(grid_ocr, text=f"{self.processor.config.raid_frame_coords['x']}, {self.processor.config.raid_frame_coords['y']}", font=ctk.CTkFont(weight="bold"))
        self.lbl_raid_coords.grid(row=0, column=1, padx=(0, 10))
        ctk.CTkButton(grid_ocr, text="Задать", command=lambda: self.open_cropper('raid'), width=80, fg_color=Theme.ACCENT_BLUE).grid(row=0, column=2)
        
        # Перс
        ctk.CTkLabel(grid_ocr, text="Координаты Статистики (X, Y):", text_color=Theme.TEXT_SECONDARY).grid(row=1, column=0, sticky="w", pady=10, padx=(0, 20))
        self.lbl_pers_coords = ctk.CTkLabel(grid_ocr, text=f"{self.processor.config.personal_frame_coords['x']}, {self.processor.config.personal_frame_coords['y']}", font=ctk.CTkFont(weight="bold"))
        self.lbl_pers_coords.grid(row=1, column=1, padx=(0, 10))
        ctk.CTkButton(grid_ocr, text="Задать", command=lambda: self.open_cropper('personal'), width=80, fg_color=Theme.ACCENT_BLUE).grid(row=1, column=2)

        # --- Алгоритмы ---
        self.create_section_header(scroll_frame, "АЛГОРИТМЫ ОБРАБОТКИ")
        card_algo = self.create_card(scroll_frame)
        card_algo.pack(fill="x", pady=10)
        grid_algo = ctk.CTkFrame(card_algo, fg_color="transparent")
        grid_algo.pack(padx=20, pady=20, fill="x")

        ctk.CTkLabel(grid_algo, text="Таймаут группы (мин):", text_color=Theme.TEXT_SECONDARY).grid(row=0, column=0, sticky="w", pady=10)
        self.var_max_diff = ctk.StringVar(value=str(self.processor.config.max_diff_time))
        entry_timeout = ctk.CTkEntry(grid_algo, textvariable=self.var_max_diff, width=150)
        entry_timeout.grid(row=0, column=1, sticky="w", padx=20)
        entry_timeout.bind("<FocusOut>", lambda e: self.save_settings())
        
        ctk.CTkLabel(grid_algo, text="Режим OCR:", text_color=Theme.TEXT_SECONDARY).grid(row=1, column=0, sticky="w", pady=10)
        current_mode = self.processor.config.get("ocr_mode")
        self.ocr_mode_map = {"offline": "Оффлайн", "mixed": "Смешанный"}
        self.ocr_mode_map_rev = {v: k for k, v in self.ocr_mode_map.items()}
        display_mode = self.ocr_mode_map.get(current_mode, "Оффлайн")
        self.var_ocr_mode = ctk.StringVar(value=display_mode)
        
        ctk.CTkComboBox(grid_algo, values=list(self.ocr_mode_map.values()), variable=self.var_ocr_mode, command=self.on_ocr_change, width=150).grid(row=1, column=1, sticky="w", padx=20)
        
        # API Key (динамический)
        self.lbl_api_key = ctk.CTkLabel(grid_algo, text="OCR.space API Key:", text_color=Theme.TEXT_SECONDARY)
        self.var_api_key = ctk.StringVar(value=self.processor.config.get("ocr_api_key"))
        self.entry_api_key = ctk.CTkEntry(grid_algo, textvariable=self.var_api_key, width=150)
        self.entry_api_key.bind("<FocusOut>", lambda e: self.save_settings())
        
        self.toggle_api_key(display_mode)
        
        # --- Отладка ---
        self.create_section_header(scroll_frame, "ПРОЧЕЕ И ОТЛАДКА")
        card_debug = self.create_card(scroll_frame)
        card_debug.pack(fill="x", pady=10)
        
        self.var_show = ctk.BooleanVar(value=self.processor.config.show_afterscreen)
        ctk.CTkCheckBox(card_debug, text="Открывать Excel после обработки", variable=self.var_show, command=self.save_settings).pack(anchor="w", padx=20, pady=(20, 10))
        
        self.var_debug_screens = ctk.BooleanVar(value=self.processor.config.get("debug_screens"))
        ctk.CTkCheckBox(card_debug, text="Сохранять отладочные скриншоты", variable=self.var_debug_screens, command=self.save_settings).pack(anchor="w", padx=20, pady=10)
        
        self.var_debug = ctk.BooleanVar(value=self.processor.config.debug)
        ctk.CTkCheckBox(card_debug, text="Режим отладки (расширенный лог)", variable=self.var_debug, command=self.save_settings).pack(anchor="w", padx=20, pady=(10, 20))

    def setup_about_view(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        
        # Hero section
        hero_card = self.create_card(parent)
        hero_card.pack(fill="x", pady=(0, 20))
        
        ctk.CTkLabel(hero_card, text="RaidStat", font=ctk.CTkFont(size=40, weight="bold"), text_color=Theme.ACCENT_BLUE).pack(pady=(30, 5))
        ctk.CTkLabel(hero_card, text="Автоматизация сбора статистики и учета посещаемости в ArcheAge", font=ctk.CTkFont(size=14), text_color=Theme.TEXT_SECONDARY).pack(pady=(0, 30))
        
        # Инфо
        info_card = self.create_card(parent)
        info_card.pack(fill="x", pady=10)
        
        ctk.CTkLabel(info_card, text="Автор: Эник", font=ctk.CTkFont(size=16)).pack(pady=(20, 5))
        ctk.CTkLabel(info_card, text="Версия: 1.0", text_color=Theme.TEXT_SECONDARY).pack(pady=(0, 20))
        
        btn_box = ctk.CTkFrame(info_card, fg_color="transparent")
        btn_box.pack(pady=10)
        
        ctk.CTkButton(
            btn_box, 
            text="📖 Инструкция (GitHub)", 
            command=lambda: webbrowser.open("https://github.com/Anykey222/raidstat/blob/main/README.md"),
            width=200
        ).pack(pady=5)
        
        self.btn_youtube = ctk.CTkButton(
            btn_box, 
            text="🎥 Видео-гайд (YouTube)", 
            command=lambda: webbrowser.open("https://youtube.com/"), 
            fg_color="#c4302b", 
            hover_color="#a82925",
            width=200
        )
        self.btn_youtube.pack(pady=5)
        
        # Donate
        donate_card = ctk.CTkFrame(parent, fg_color="#0ea5e9", corner_radius=15) # Sky blue
        donate_card.pack(fill="x", pady=20)
        
        ctk.CTkLabel(donate_card, text="❤️ Поддержать проект", font=ctk.CTkFont(size=18, weight="bold"), text_color="white").pack(pady=(15, 5))
        ctk.CTkLabel(donate_card, text="Ваша поддержка помогает проекту развиваться быстрее.", text_color="white").pack(pady=5)
        
        ctk.CTkButton(
            donate_card, 
            text="Отправить донат (CloudTips)", 
            command=lambda: webbrowser.open("https://pay.cloudtips.ru/p/17913652"),
            fg_color="white", 
            text_color="#0284c7",
            hover_color="#f0f9ff", 
            font=ctk.CTkFont(weight="bold")
        ).pack(pady=(10, 20))

    # --- Logic ---

    def _handle_console_keypress(self, event):
        ctrl_pressed = (event.state & 0x4) != 0
        if not ctrl_pressed:
            return
        if event.keycode == 65:  # A
            return self.select_all_log(event)
        elif event.keycode == 67:  # C
            return self.copy_selection_log(event)

    def select_all_log(self, event):
        try:
            self.console_text._textbox.tag_add("sel", "1.0", "end")
        except Exception:
            pass
        return "break"

    def copy_selection_log(self, event):
        try:
            text = self.console_text._textbox.get("sel.first", "sel.last")
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()
        except tk.TclError:
            pass 
        except Exception as e:
            logging.error(f"Ошибка копирования: {e}")
        return "break"

    def on_ocr_change(self, choice):
        self.toggle_api_key(choice)
        self.save_settings()

    def toggle_api_key(self, mode_display):
        if "Смешанный" in mode_display:
            self.lbl_api_key.grid(row=2, column=0, sticky="w", pady=10)
            self.entry_api_key.grid(row=2, column=1, sticky="w", padx=20, pady=10)
        else:
            self.lbl_api_key.grid_forget()
            self.entry_api_key.grid_forget()

    def select_folder(self, var):
        folder = filedialog.askdirectory()
        if folder:
            var.set(folder)
            self.processor.config.set("screenshots_directory", folder)
            # Sync
            if var == self.att_folder_path:
                self.stat_folder_path.set(folder)
            else:
                self.att_folder_path.set(folder)

    def save_attendance_options(self):
        self.processor.config.set("recursive_scan", self.att_recursive.get())
        logging.info(f"Настройка 'Обрабатывать подпапки' сохранена: {self.att_recursive.get()}")

    def save_settings(self, *args):
        try:
            val_scale = int(self.var_scale.get())
            self.processor.config.set("interface_scale", val_scale)
            # Примечание: масштаб обычно требует перезапуска, но сохраняем
            
            try:
                self.processor.config.set("max_diff_time", int(self.var_max_diff.get()))
            except ValueError:
                 logging.warning("Некорректное значение для таймаута группы.")
            
            self.processor.config.set("show_afterscreen", self.var_show.get())
            self.processor.config.set("debug_screens", self.var_debug_screens.get())
            self.processor.config.set("debug", self.var_debug.get())
            
            logging.getLogger().setLevel(logging.DEBUG if self.var_debug.get() else logging.INFO)
            
            mode_display = self.var_ocr_mode.get()
            mode_internal = self.ocr_mode_map_rev.get(mode_display, "offline")
            
            self.processor.config.set("ocr_mode", mode_internal)
            self.processor.config.set("ocr_api_key", self.var_api_key.get())
            self.processor.reload_config()
            
            logging.info(f"Настройки сохранены.")
        except Exception as e:
            logging.error(f"Ошибка сохранения настроек: {e}")

    def open_cropper(self, mode):
        img_path = filedialog.askopenfilename(filetypes=[("Images", "*.jpg *.png *.bmp")])
        if not img_path:
            return
            
        def callback(x, y):
            if mode == 'raid':
                self.processor.config.set("raid_frame_coords", {"x": x, "y": y})
                self.lbl_raid_coords.configure(text=f"{x}, {y}")
            else:
                self.processor.config.set("personal_frame_coords", {"x": x, "y": y})
                self.lbl_pers_coords.configure(text=f"{x}, {y}")
            self.processor.reload_config()
            logging.info(f"Обновлены координаты {mode}: {x}, {y}")

        CropWindow(self, img_path, callback=callback)

    def stop_processing_action(self):
        self.processor.stop_processing()

    def run_attendance(self):
        path = self.att_folder_path.get()
        if not path or path == "Папка не выбрана":
            logging.warning("Выберите папку со скриншотами.")
            return
            
        self.btn_att_process.configure(text="🛑 Остановить", fg_color=Theme.ACCENT_RED, hover_color=Theme.BTN_HOVER_RED, command=self.stop_processing_action)
        
        def task():
            try:
                count = self.processor.process_attendance(path, self.att_recursive.get())
                if self.processor.stop_event.is_set():
                    logging.info("Обработка остановлена пользователем.")
                else:
                    logging.info(f"Обработка посещаемости завершена. Найдено имен: {count}.")
                    
                    excel_path = self.processor.storage.file_path
                    if os.path.exists(excel_path):
                        logging.info(f"Открываю {excel_path}...")
                        os.startfile(excel_path)
                    else:
                        logging.warning(f"Файл Excel не найден: {excel_path}")
            except Exception as e:
                logging.error(f"Ошибка: {e}")
                import traceback
                logging.error(traceback.format_exc())
            finally:
                self.after(0, lambda: self.btn_att_process.configure(
                    text="🚀 Начать обработку", fg_color=Theme.ACCENT_GREEN, hover_color=Theme.BTN_HOVER_GREEN, command=self.run_attendance
                ))
        
        threading.Thread(target=task).start()

    def run_statistics(self):
        path = self.stat_folder_path.get()
        if not path or path == "Папка не выбрана":
            logging.warning("Выберите папку со скриншотами.")
            return
            
        self.btn_stat_process.configure(text="🛑 Остановить", fg_color=Theme.ACCENT_RED, hover_color=Theme.BTN_HOVER_RED, command=self.stop_processing_action)
        
        def task():
            try:
                count = self.processor.process_statistics(path, recursive=False)
                if self.processor.stop_event.is_set():
                     logging.info("Сбор статистики остановлен пользователем.")
                else:
                    if count == 0:
                        logging.warning(f"Изображений в папке {path} не найдено.")
                    else:
                        logging.info(f"Сбор статистики завершен. Обработано изображений: {count}.")
                        
                        excel_path = self.processor.storage.file_path
                        if os.path.exists(excel_path):
                            logging.info(f"Открываю {excel_path}...")
                            os.startfile(excel_path)
                        else:
                            logging.warning(f"Файл Excel не найден: {excel_path}")
            except Exception as e:
                logging.error(f"Ошибка: {e}")
                import traceback
                logging.error(traceback.format_exc())
            finally:
                self.after(0, lambda: self.btn_stat_process.configure(
                    text="📊 Собрать статистику", fg_color=Theme.ACCENT_GREEN, hover_color=Theme.BTN_HOVER_GREEN, command=self.run_statistics
                ))
        
        threading.Thread(target=task).start()

    def revert_attendance_ui(self):
        def task():
            try:
                self.processor.revert_attendance()
                logging.info("Скриншоты посещаемости возвращены (reverted).")
            except Exception as e:
                logging.error(f"Ошибка при возврате скриншотов: {e}")
        threading.Thread(target=task).start()

    def revert_statistics_ui(self):
        def task():
            try:
                self.processor.revert_statistics()
                logging.info("Скриншоты статистики возвращены (reverted).")
            except Exception as e:
                logging.error(f"Ошибка при возврате скриншотов: {e}")
        threading.Thread(target=task).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()
