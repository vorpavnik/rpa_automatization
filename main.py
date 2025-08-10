import locale
import tkinter as tk
from datetime import datetime
from tkinter import scrolledtext
import threading
import logging
import queue
import sys
import io
from logging.handlers import QueueHandler
from gui_app import primary_work
from const import IMAGE_PATH, APP_TITLE, APP_SIZE, LOG_FILE

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
locale.setlocale(locale.LC_ALL, '')


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_SIZE)
        self.IMAGE_PATH = IMAGE_PATH

        # Настройка логгирования
        self.setup_logging()
        self.create_widgets()
        self.check_log_queue()

    def setup_logging(self):
        """Настройка двойного логгирования (файл + UI)"""
        # Лог в файл
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            handlers=[logging.FileHandler(LOG_FILE)]
        )

        # Лог в UI
        self.log_queue = queue.Queue()
        queue_handler = QueueHandler(self.log_queue)
        ui_logger = logging.getLogger('UI')
        ui_logger.setLevel(logging.INFO)
        ui_logger.addHandler(queue_handler)

    def create_widgets(self):
        """Создание интерфейса с увеличенными кнопками"""
        # Заголовок
        title_label = tk.Label(
            self.root,
            text=APP_TITLE,
            font=("Arial", 16),
            pady=15
        )
        title_label.pack()

        # Лог
        self.log_area = scrolledtext.ScrolledText(
            self.root,
            width=100,
            height=25,
            font=("Courier", 10),
            wrap=tk.WORD
        )
        self.log_area.pack(padx=15, pady=10, expand=True, fill=tk.BOTH)

        # Кнопки
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=15)

        start_button = tk.Button(
            button_frame,
            text="Запустить ClientManager",
            font=("Arial", 12),
            width=25,
            height=2,
            command=self.start_task
        )
        start_button.pack(side=tk.LEFT, padx=20)

        close_button = tk.Button(
            button_frame,
            text="Закрыть программу",
            font=("Arial", 12),
            width=25,
            height=2,
            command=self.root.quit
        )
        close_button.pack(side=tk.LEFT, padx=20)

    def log(self, message):
        """Логирование в UI и файл"""
        timestamped_msg = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"

        # В UI
        self.log_area.insert(tk.END, timestamped_msg + "\n")
        self.log_area.see(tk.END)

        # В файл
        logging.info(message)

    def check_log_queue(self):
        """Проверка очереди логов"""
        while not self.log_queue.empty():
            record = self.log_queue.get()
            self.log_area.insert(tk.END, record.getMessage() + "\n")
            self.log_area.see(tk.END)
        self.root.after(100, self.check_log_queue)

    def start_task(self):
        """Запуск задачи в отдельном потоке"""
        self.log(f"▶️ Начало взаимодействия с графиком МВПС.")
        thread = threading.Thread(
            target=primary_work,
            args=(self,),
            daemon=True
        )
        thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
