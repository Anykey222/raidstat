"""
Модуль для визуального выбора координат на скриншоте.
Ctrl + колесико мыши - зум в позицию курсора.
Просто колесико - прокрутка вверх/вниз.
Shift + колесико - прокрутка влево/вправо.
Клик левой кнопкой - выбор точки.

Оптимизация: рендерится только видимая область изображения.
"""

import customtkinter as ctk
from PIL import Image, ImageTk
import tkinter as tk


class CropWindow(ctk.CTkToplevel):
    def __init__(self, master, image_path, initial_coords=None, callback=None):
        super().__init__(master)
        self.title("Выбор области - Ctrl+Колесо=Зум, Колесо=Прокрутка")
        self.geometry("1200x800")
        
        self.callback = callback
        self.image_path = image_path
        
        # Загрузка оригинального изображения
        self.pil_image = Image.open(image_path)
        self.img_width = self.pil_image.width
        self.img_height = self.pil_image.height
        
        # Масштаб
        self.current_scale = 1.0
        self.min_scale = 0.1
        self.max_scale = 20.0
        
        # Текущая позиция просмотра (левый верхний угол в реальных координатах)
        self.view_x = 0.0
        self.view_y = 0.0
        
        # Контейнер для canvas
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(fill="both", expand=True)
        
        # Canvas без scrollbars - управляем видом напрямую
        self.canvas = tk.Canvas(
            self.frame, 
            cursor="cross", 
            bg="#2b2b2b", 
            highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True)
        
        # Изображение на canvas
        self.tk_image = None
        self.image_id = None
        
        # Координаты для перекрестия (в реальных координатах)
        self._crosshair_coords = None
        
        # Переменные для перетаскивания (ПКМ или средняя кнопка)
        self.drag_data = {"x": 0, "y": 0, "dragging": False}
        
        # Привязка событий
        self.canvas.bind("<ButtonPress-1>", self.on_left_click)
        self.canvas.bind("<ButtonPress-2>", self.on_drag_start)  # Средняя кнопка
        self.canvas.bind("<ButtonPress-3>", self.on_drag_start)  # Правая кнопка
        self.canvas.bind("<B2-Motion>", self.on_drag_move)
        self.canvas.bind("<B3-Motion>", self.on_drag_move)
        self.canvas.bind("<ButtonRelease-2>", self.on_drag_end)
        self.canvas.bind("<ButtonRelease-3>", self.on_drag_end)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Control-MouseWheel>", self.on_ctrl_mouse_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self.on_shift_mouse_wheel)
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        
        # Начальные координаты для перекрестия
        if initial_coords:
            self._crosshair_coords = (initial_coords.get('x', 0), initial_coords.get('y', 0))
            # Центрируем на этих координатах после отрисовки
            self.after(100, lambda: self._center_on_point(*self._crosshair_coords))
        
        # Первая отрисовка после того как окно появится
        self.after(50, self._render_view)
        
        # Окно поверх основного и захват фокуса
        self.transient(master)
        self.grab_set()
        self.focus_force()
        self.lift()
        self.after(100, self.lift)
    
    def _render_view(self):
        """Рендерит только видимую часть изображения - быстро!"""
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        if canvas_w <= 1 or canvas_h <= 1:
            # Canvas ещё не готов
            self.after(50, self._render_view)
            return
        
        # Размер видимой области в реальных координатах
        view_w_real = canvas_w / self.current_scale
        view_h_real = canvas_h / self.current_scale
        
        # Ограничиваем view_x, view_y
        max_view_x = max(0, self.img_width - view_w_real)
        max_view_y = max(0, self.img_height - view_h_real)
        self.view_x = max(0, min(self.view_x, max_view_x))
        self.view_y = max(0, min(self.view_y, max_view_y))
        
        # Координаты кропа в оригинальном изображении
        crop_x1 = int(self.view_x)
        crop_y1 = int(self.view_y)
        crop_x2 = min(self.img_width, int(self.view_x + view_w_real) + 1)
        crop_y2 = min(self.img_height, int(self.view_y + view_h_real) + 1)
        
        # Кропаем только нужную часть
        cropped = self.pil_image.crop((crop_x1, crop_y1, crop_x2, crop_y2))
        
        # Масштабируем до размера canvas
        target_w = int((crop_x2 - crop_x1) * self.current_scale)
        target_h = int((crop_y2 - crop_y1) * self.current_scale)
        
        if target_w > 0 and target_h > 0:
            # NEAREST - самый быстрый, при масштабе > 2 виден пиксельный эффект (что удобно для точности)
            if self.current_scale >= 2.0:
                resized = cropped.resize((target_w, target_h), Image.Resampling.NEAREST)
            else:
                resized = cropped.resize((target_w, target_h), Image.Resampling.BILINEAR)
            
            self.tk_image = ImageTk.PhotoImage(resized)
            
            # Смещение отображения (для дробных координат)
            offset_x = -(self.view_x - crop_x1) * self.current_scale
            offset_y = -(self.view_y - crop_y1) * self.current_scale
            
            if self.image_id is None:
                self.image_id = self.canvas.create_image(offset_x, offset_y, image=self.tk_image, anchor="nw")
            else:
                self.canvas.coords(self.image_id, offset_x, offset_y)
                self.canvas.itemconfig(self.image_id, image=self.tk_image)
        
        # Перерисовываем перекрестие
        self._draw_crosshair()
    
    def _draw_crosshair(self):
        """Рисует перекрестие если есть координаты."""
        self.canvas.delete("crosshair")
        
        if self._crosshair_coords is None:
            return
        
        real_x, real_y = self._crosshair_coords
        
        # Координаты на canvas
        canvas_x = (real_x - self.view_x) * self.current_scale
        canvas_y = (real_y - self.view_y) * self.current_scale
        
        # Проверяем видимость
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        if canvas_x < -20 or canvas_x > canvas_w + 20 or canvas_y < -20 or canvas_y > canvas_h + 20:
            return  # Не видно
        
        # Рисуем
        size = 15
        self.canvas.create_line(canvas_x - size, canvas_y, canvas_x + size, canvas_y, 
                                fill="red", width=2, tags="crosshair")
        self.canvas.create_line(canvas_x, canvas_y - size, canvas_x, canvas_y + size, 
                                fill="red", width=2, tags="crosshair")
        self.canvas.create_oval(canvas_x - 4, canvas_y - 4, canvas_x + 4, canvas_y + 4,
                                fill="red", outline="white", width=1, tags="crosshair")
    
    def _center_on_point(self, real_x, real_y):
        """Центрирует вид на указанной точке."""
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        view_w_real = canvas_w / self.current_scale
        view_h_real = canvas_h / self.current_scale
        
        self.view_x = real_x - view_w_real / 2
        self.view_y = real_y - view_h_real / 2
        
        self._render_view()
    
    def on_canvas_resize(self, event):
        """При изменении размера окна перерисовываем."""
        self._render_view()
    
    def on_mouse_wheel(self, event):
        """Прокрутка вверх/вниз."""
        if event.state & 0x4:  # Ctrl зажат - игнорируем
            return
        
        # Скорость прокрутки зависит от масштаба
        scroll_speed = 50 / self.current_scale
        
        if event.delta > 0:
            self.view_y -= scroll_speed
        else:
            self.view_y += scroll_speed
        
        self._render_view()
    
    def on_shift_mouse_wheel(self, event):
        """Прокрутка влево/вправо."""
        scroll_speed = 50 / self.current_scale
        
        if event.delta > 0:
            self.view_x -= scroll_speed
        else:
            self.view_x += scroll_speed
        
        self._render_view()
    
    def on_ctrl_mouse_wheel(self, event):
        """Зум с центрированием на позиции курсора."""
        # Позиция курсора на canvas
        mouse_x = event.x
        mouse_y = event.y
        
        # Реальные координаты под курсором
        real_x = self.view_x + mouse_x / self.current_scale
        real_y = self.view_y + mouse_y / self.current_scale
        
        # Новый масштаб
        if event.delta > 0:
            new_scale = self.current_scale * 1.25
        else:
            new_scale = self.current_scale / 1.25
        
        new_scale = max(self.min_scale, min(self.max_scale, new_scale))
        
        if abs(new_scale - self.current_scale) < 0.001:
            return
        
        self.current_scale = new_scale
        
        # Пересчитываем view_x, view_y чтобы точка под курсором осталась на месте
        self.view_x = real_x - mouse_x / self.current_scale
        self.view_y = real_y - mouse_y / self.current_scale
        
        self._render_view()
    
    def on_drag_start(self, event):
        """Начало перетаскивания."""
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        self.drag_data["dragging"] = True
        self.canvas.config(cursor="fleur")
    
    def on_drag_move(self, event):
        """Перетаскивание."""
        if not self.drag_data["dragging"]:
            return
        
        dx = self.drag_data["x"] - event.x
        dy = self.drag_data["y"] - event.y
        
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        
        # Смещаем вид
        self.view_x += dx / self.current_scale
        self.view_y += dy / self.current_scale
        
        self._render_view()
    
    def on_drag_end(self, event):
        """Конец перетаскивания."""
        self.drag_data["dragging"] = False
        self.canvas.config(cursor="cross")
    
    def on_left_click(self, event):
        """Выбор точки."""
        # Реальные координаты
        real_x = int(self.view_x + event.x / self.current_scale)
        real_y = int(self.view_y + event.y / self.current_scale)
        
        # Проверяем границы
        if real_x < 0 or real_x >= self.img_width or real_y < 0 or real_y >= self.img_height:
            return
        
        # Сохраняем и рисуем
        self._crosshair_coords = (real_x, real_y)
        self._draw_crosshair()
        
        # Callback и закрытие
        if self.callback:
            self.callback(real_x, real_y)
        
        self.grab_release()
        self.destroy()
