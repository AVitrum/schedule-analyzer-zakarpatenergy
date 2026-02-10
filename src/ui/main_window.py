"""
Main application window for the Schedule Analyzer.

Redesigned with improved user experience and modern interface.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image
import json
from datetime import datetime
import os
import sys
from typing import Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from analyze_schedule import analyze_row, time_to_string, resize_image
from src.themes.theme_manager import ModernTheme
from src.utils.calendar_export import CalendarExporter
from src.utils.time_calculator import calculate_total_time, queue_sort_key


class ModernStyleConfigurator:
    """Enhanced style configurator with modern design elements."""

    def __init__(self, colors: Dict[str, str]):
        self.colors = colors
        self.style = ttk.Style()

    def configure_all_styles(self, root: tk.Tk) -> None:
        self.style.theme_use('clam')
        root.configure(bg=self.colors['bg'])

        self._configure_frames()
        self._configure_labels()
        self._configure_inputs()
        self._configure_buttons()

    def _configure_frames(self) -> None:
        self.style.configure('Modern.TFrame', background=self.colors['bg'])

        self.style.configure('Card.TFrame',
                           background=self.colors['secondary_bg'],
                           relief='solid',
                           borderwidth=1)

    def _configure_labels(self) -> None:
        self.style.configure('Modern.TLabel',
                           background=self.colors['bg'],
                           foreground=self.colors['fg'],
                           font=('Segoe UI', 10))

        self.style.configure('Title.TLabel',
                           background=self.colors['bg'],
                           foreground=self.colors['fg'],
                           font=('Segoe UI', 24, 'bold'))

        self.style.configure('Heading.TLabel',
                           background=self.colors['secondary_bg'],
                           foreground=self.colors['fg'],
                           font=('Segoe UI', 14, 'bold'))

        self.style.configure('Subtitle.TLabel',
                           background=self.colors['bg'],
                           foreground=self.colors['text_secondary'],
                           font=('Segoe UI', 11))

        self.style.configure('Card.TLabel',
                           background=self.colors['secondary_bg'],
                           foreground=self.colors['fg'],
                           font=('Segoe UI', 10))

        self.style.configure('Info.TLabel',
                           background=self.colors['secondary_bg'],
                           foreground=self.colors['text_secondary'],
                           font=('Segoe UI', 9))

    def _configure_inputs(self) -> None:
        self.style.configure('Modern.TEntry',
                           fieldbackground=self.colors['input_bg'],
                           foreground=self.colors['input_fg'],
                           insertcolor=self.colors['input_fg'],
                           borderwidth=2,
                           relief='solid')

        self.style.map('Modern.TEntry',
                     fieldbackground=[('readonly', self.colors['input_bg']),
                                    ('disabled', self.colors['input_bg'])],
                     foreground=[('readonly', self.colors['input_fg']),
                               ('disabled', self.colors['text_secondary'])],
                     bordercolor=[('focus', self.colors['accent'])])

        self.style.configure('Modern.TCombobox',
                           fieldbackground=self.colors['input_bg'],
                           background=self.colors['input_bg'],
                           foreground=self.colors['input_fg'],
                           selectbackground=self.colors['accent'],
                           selectforeground=self.colors['button_fg'],
                           borderwidth=2,
                           arrowcolor=self.colors['fg'],
                           insertcolor=self.colors['input_fg'])

        self.style.map('Modern.TCombobox',
                     fieldbackground=[('readonly', self.colors['input_bg'])],
                     bordercolor=[('focus', self.colors['accent'])])

    def _configure_buttons(self) -> None:
        self.style.configure('Primary.TButton',
                           background=self.colors['accent'],
                           foreground=self.colors['button_fg'],
                           borderwidth=0,
                           focuscolor='none',
                           font=('Segoe UI', 11, 'bold'),
                           relief='flat',
                           padding=(20, 12))

        self.style.map('Primary.TButton',
                     background=[('active', self.colors['button_hover']),
                               ('pressed', self.colors['accent'])],
                     foreground=[('active', self.colors['button_fg'])])

        self.style.configure('Secondary.TButton',
                           background=self.colors['secondary_bg'],
                           foreground=self.colors['fg'],
                           borderwidth=2,
                           focuscolor='none',
                           font=('Segoe UI', 10),
                           relief='solid',
                           padding=(15, 8))

        self.style.map('Secondary.TButton',
                     background=[('active', self.colors['input_bg'])],
                     bordercolor=[('active', self.colors['accent'])])

        self.style.configure('Icon.TButton',
                           background=self.colors['secondary_bg'],
                           foreground=self.colors['accent'],
                           borderwidth=0,
                           focuscolor='none',
                           font=('Segoe UI', 10, 'bold'),
                           relief='flat',
                           padding=(10, 8))

        self.style.map('Icon.TButton',
                     background=[('active', self.colors['input_bg'])])


class ScheduleAnalyzerUI:
    """Main application window with modern, user-friendly interface."""

    WINDOW_TITLE = "Аналізатор Графіка Відключень"
    WINDOW_SIZE = "1000x700"
    MIN_WINDOW_SIZE = (900, 600)

    AVAILABLE_QUEUES = {
        "Черга 1-1": 90, "Черга 1-2": 109,
        "Черга 2-1": 127, "Черга 2-2": 146,
        "Черга 3-1": 170, "Черга 3-2": 190,
        "Черга 4-1": 214, "Черга 4-2": 233,
        "Черга 5-1": 255, "Черга 5-2": 275,
        "Черга 6-1": 298, "Черга 6-2": 315,
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(self.WINDOW_TITLE)
        self.root.geometry(self.WINDOW_SIZE)
        self.root.minsize(*self.MIN_WINDOW_SIZE)

        self.dark_mode = ModernTheme.detect_system_theme()
        self.colors = ModernTheme.get_theme_colors(self.dark_mode)

        self.setup_styles()
        self.center_window()
        self.initialize_variables()
        self.create_modern_ui()

    def setup_styles(self) -> None:
        configurator = ModernStyleConfigurator(self.colors)
        configurator.configure_all_styles(self.root)

    def center_window(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def initialize_variables(self) -> None:
        self.image_path = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.selected_queue = tk.StringVar(value="Черга 1-1")
        self.status_var = tk.StringVar(value="Готовий до роботи")
        self.file_name_var = tk.StringVar(value="Файл не вибрано")

    def create_modern_ui(self) -> None:
        """Create modern, simplified UI with better visual hierarchy."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_container = ttk.Frame(self.root, style='Modern.TFrame')
        main_container.grid(row=0, column=0, sticky='nsew')
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(1, weight=1)

        self._create_header_section(main_container)
        self._create_content_section(main_container)
        self._create_status_bar(main_container)

    def _create_header_section(self, parent: ttk.Frame) -> None:
        """Create compact header with title and quick info."""
        header_frame = ttk.Frame(parent, style='Modern.TFrame')
        header_frame.grid(row=0, column=0, sticky='ew', padx=20, pady=(20, 10))
        header_frame.columnconfigure(0, weight=1)

        title_label = ttk.Label(header_frame, text="Аналізатор Графіка",
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, sticky='w')

        subtitle_label = ttk.Label(header_frame,
                                   text="Швидкий аналіз та експорт в календар",
                                   style='Subtitle.TLabel')
        subtitle_label.grid(row=1, column=0, sticky='w', pady=(5, 0))

    def _create_content_section(self, parent: ttk.Frame) -> None:
        """Create main content area with two columns."""
        content_frame = ttk.Frame(parent, style='Modern.TFrame')
        content_frame.grid(row=1, column=0, sticky='nsew', padx=20, pady=10)
        content_frame.columnconfigure(0, weight=1)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)

        self._create_left_panel(content_frame)
        self._create_right_panel(content_frame)

    def _create_left_panel(self, parent: ttk.Frame) -> None:
        """Create left panel with inputs and actions."""
        left_panel = ttk.Frame(parent, style='Modern.TFrame')
        left_panel.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        left_panel.columnconfigure(0, weight=1)

        self._create_input_card(left_panel)
        self._create_action_card(left_panel)

    def _create_input_card(self, parent: ttk.Frame) -> None:
        """Create modern input card with file, date, and queue selection."""
        card = tk.Frame(parent, bg=self.colors['secondary_bg'],
                       highlightbackground=self.colors['border'],
                       highlightthickness=1)
        card.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        card.columnconfigure(0, weight=1)

        header = ttk.Label(card, text="Параметри аналізу", style='Heading.TLabel')
        header.grid(row=0, column=0, sticky='w', padx=20, pady=(20, 15))

        file_frame = ttk.Frame(card, style='Card.TFrame')
        file_frame.grid(row=1, column=0, sticky='ew', padx=20, pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)

        file_info_frame = ttk.Frame(file_frame, style='Card.TFrame')
        file_info_frame.grid(row=0, column=0, sticky='ew')
        file_info_frame.columnconfigure(0, weight=1)

        ttk.Label(file_info_frame, text="Графік відключень",
                 style='Card.TLabel', font=('Segoe UI', 10, 'bold')).grid(
                     row=0, column=0, sticky='w')

        file_name_label = ttk.Label(file_info_frame, textvariable=self.file_name_var,
                                    style='Info.TLabel')
        file_name_label.grid(row=1, column=0, sticky='w', pady=(2, 0))

        select_btn = ttk.Button(file_frame, text="Вибрати файл",
                               command=self.select_file_modern,
                               style='Icon.TButton')
        select_btn.grid(row=0, column=1, padx=(10, 0))

        separator1 = ttk.Separator(card, orient='horizontal')
        separator1.grid(row=2, column=0, sticky='ew', padx=20, pady=10)

        date_frame = ttk.Frame(card, style='Card.TFrame')
        date_frame.grid(row=3, column=0, sticky='ew', padx=20, pady=(0, 15))
        date_frame.columnconfigure(1, weight=1)

        ttk.Label(date_frame, text="Дата:", style='Card.TLabel',
                 font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w')

        date_entry = ttk.Entry(date_frame, textvariable=self.date_var,
                              font=('Segoe UI', 11), style='Modern.TEntry',
                              width=15)
        date_entry.grid(row=0, column=1, sticky='w', padx=(10, 0), ipady=6)

        ttk.Label(date_frame, text="РРРР-ММ-ДД", style='Info.TLabel').grid(
            row=0, column=2, padx=(10, 0))

        separator2 = ttk.Separator(card, orient='horizontal')
        separator2.grid(row=4, column=0, sticky='ew', padx=20, pady=10)

        queue_frame = ttk.Frame(card, style='Card.TFrame')
        queue_frame.grid(row=5, column=0, sticky='ew', padx=20, pady=(0, 20))
        queue_frame.columnconfigure(1, weight=1)

        ttk.Label(queue_frame, text="Черга:", style='Card.TLabel',
                 font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, sticky='w')

        queue_combo = ttk.Combobox(queue_frame, textvariable=self.selected_queue,
                                   values=list(self.AVAILABLE_QUEUES.keys()),
                                   state='readonly', font=('Segoe UI', 11),
                                   style='Modern.TCombobox', width=18)
        queue_combo.grid(row=0, column=1, sticky='ew', padx=(10, 0), ipady=6)

    def _create_action_card(self, parent: ttk.Frame) -> None:
        """Create action buttons card."""
        card = tk.Frame(parent, bg=self.colors['secondary_bg'],
                       highlightbackground=self.colors['border'],
                       highlightthickness=1)
        card.grid(row=1, column=0, sticky='ew', pady=(0, 0))
        card.columnconfigure(0, weight=1)

        header = ttk.Label(card, text="Дії", style='Heading.TLabel')
        header.grid(row=0, column=0, sticky='w', padx=20, pady=(20, 15))

        analyze_btn = ttk.Button(card, text="Аналізувати графік",
                                command=self.analyze_schedule_modern,
                                style='Primary.TButton')
        analyze_btn.grid(row=1, column=0, sticky='ew', padx=20, pady=(0, 10))

        compare_btn = ttk.Button(card, text="Порівняти всі черги",
                                command=self.compare_all_queues,
                                style='Secondary.TButton')
        compare_btn.grid(row=2, column=0, sticky='ew', padx=20, pady=(0, 20))

    def _create_right_panel(self, parent: ttk.Frame) -> None:
        """Create right panel with results."""
        right_panel = ttk.Frame(parent, style='Modern.TFrame')
        right_panel.grid(row=0, column=1, sticky='nsew', padx=(10, 0))
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(0, weight=1)

        card = tk.Frame(right_panel, bg=self.colors['secondary_bg'],
                       highlightbackground=self.colors['border'],
                       highlightthickness=1)
        card.grid(row=0, column=0, sticky='nsew')
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        header_frame = ttk.Frame(card, style='Card.TFrame')
        header_frame.grid(row=0, column=0, sticky='ew', padx=20, pady=(20, 10))
        header_frame.columnconfigure(0, weight=1)

        ttk.Label(header_frame, text="Результати", style='Heading.TLabel').grid(
            row=0, column=0, sticky='w')

        self.result_text = scrolledtext.ScrolledText(
            card,
            wrap=tk.WORD,
            font=('Consolas', 10),
            bg=self.colors['input_bg'],
            fg=self.colors['input_fg'],
            relief=tk.FLAT,
            borderwidth=0,
            insertbackground=self.colors['fg'],
            padx=10,
            pady=10
        )
        self.result_text.grid(row=1, column=0, sticky='nsew', padx=20, pady=(0, 10))

        button_frame = ttk.Frame(card, style='Card.TFrame')
        button_frame.grid(row=2, column=0, sticky='ew', padx=20, pady=(0, 20))

        ttk.Button(button_frame, text="Копіювати",
                  command=self.copy_result, style='Icon.TButton').grid(
                      row=0, column=0, padx=(0, 5))

        ttk.Button(button_frame, text="Зберегти",
                  command=self.save_result, style='Icon.TButton').grid(
                      row=0, column=1, padx=5)

        ttk.Button(button_frame, text="Календар",
                  command=self.export_calendar, style='Icon.TButton').grid(
                      row=0, column=2, padx=5)

        ttk.Button(button_frame, text="Очистити",
                  command=self.clear_result, style='Icon.TButton').grid(
                      row=0, column=3, padx=(5, 0))

    def _create_status_bar(self, parent: ttk.Frame) -> None:
        """Create status bar at bottom."""
        status_frame = ttk.Frame(parent, style='Modern.TFrame')
        status_frame.grid(row=2, column=0, sticky='ew', padx=20, pady=(10, 15))

        status_label = ttk.Label(status_frame, textvariable=self.status_var,
                                style='Info.TLabel')
        status_label.pack(side=tk.LEFT)

        version_label = ttk.Label(status_frame, text="v2.0.0", style='Info.TLabel')
        version_label.pack(side=tk.RIGHT)

    def select_file_modern(self) -> None:
        """Modern file selection with improved feedback."""
        filename = filedialog.askopenfilename(
            title="Виберіть файл графіку",
            filetypes=[
                ("Зображення", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("Всі файли", "*.*")
            ]
        )
        if filename:
            self.image_path.set(filename)
            short_name = os.path.basename(filename)
            if len(short_name) > 35:
                short_name = short_name[:32] + "..."
            self.file_name_var.set(short_name)
            self.status_var.set(f"Файл вибрано: {short_name}")

    def validate_inputs(self) -> Optional[datetime]:
        """Validate user inputs before analysis."""
        if not self.image_path.get():
            messagebox.showerror("Помилка", "Будь ласка, виберіть файл графіку!")
            self.status_var.set("Помилка: файл не вибрано")
            return None

        if not self.date_var.get():
            messagebox.showerror("Помилка", "Будь ласка, вкажіть дату!")
            self.status_var.set("Помилка: дата не вказана")
            return None

        try:
            date_obj = datetime.strptime(self.date_var.get(), '%Y-%m-%d')
            self.status_var.set("Дані валідні")
            return date_obj
        except ValueError:
            messagebox.showerror("Помилка",
                               "Невірний формат дати!\nВикористовуйте формат: РРРР-ММ-ДД")
            self.status_var.set("Помилка: невірний формат дати")
            return None

    def analyze_schedule_modern(self) -> None:
        """Analyze schedule with improved user feedback."""
        date_obj = self.validate_inputs()
        if not date_obj:
            return

        if not self.selected_queue.get():
            messagebox.showerror("Помилка", "Будь ласка, виберіть чергу для аналізу!")
            return

        self.status_var.set("Аналізую графік...")
        self.root.update_idletasks()

        try:
            img = Image.open(self.image_path.get())
            img = resize_image(img)

            queue_name = self.selected_queue.get()
            y_coord = self.AVAILABLE_QUEUES[queue_name]
            outages = analyze_row(img, y_coord)

            outage_list = [
                {"start": time_to_string(start), "end": time_to_string(end)}
                for start, end in outages
            ]

            hours, minutes, total_minutes = calculate_total_time(outage_list)

            output = {
                "date": date_obj.strftime('%d.%m.%Y'),
                "queue": queue_name,
                "outages": outage_list,
                "total_outage_time": {
                    "hours": hours,
                    "minutes": minutes,
                    "total_minutes": total_minutes,
                    "formatted": f"{hours} год {minutes} хв"
                },
                "analysis_info": {
                    "outage_count": len(outage_list),
                    "analyzed_at": datetime.now().strftime('%d.%m.%Y %H:%M:%S')
                }
            }

            json_output = json.dumps(output, ensure_ascii=False, indent=2)
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', json_output)

            self.status_var.set(f"Аналіз завершено: {len(outage_list)} відключень, {hours}г {minutes}хв")

            messagebox.showinfo("Успіх",
                f"Аналіз завершено успішно!\n\n"
                f"Відключень знайдено: {len(outage_list)}\n"
                f"Загальний час: {hours} год {minutes} хв\n"
                f"Дата: {date_obj.strftime('%d.%m.%Y')}")

        except FileNotFoundError:
            messagebox.showerror("Помилка", "Файл не знайдено!")
            self.status_var.set("Помилка: файл не знайдено")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка під час аналізу:\n{str(e)}")
            self.status_var.set(f"Помилка: {str(e)[:50]}")

    def compare_all_queues(self) -> None:
        """Compare all queues with modern visualization."""
        date_obj = self.validate_inputs()
        if not date_obj:
            return

        self.status_var.set("Порівнюю всі черги...")
        self.root.update_idletasks()

        try:
            img = Image.open(self.image_path.get())
            img = resize_image(img)

            comparison_results = self._analyze_all_queues(img)
            comparison_results.sort(key=queue_sort_key)

            self._show_modern_comparison_window(comparison_results, date_obj.strftime('%d.%m.%Y'))
            self.status_var.set(f"Порівняння завершено для {len(comparison_results)} черг")

        except FileNotFoundError:
            messagebox.showerror("Помилка", "Файл не знайдено!")
            self.status_var.set("Помилка: файл не знайдено")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка під час порівняння:\n{str(e)}")
            self.status_var.set(f"Помилка: {str(e)[:50]}")

    def _analyze_all_queues(self, img: Image.Image) -> List[Dict]:
        """Analyze all queues in the image."""
        comparison_results = []

        for queue_name, y_coord in self.AVAILABLE_QUEUES.items():
            outages = analyze_row(img, y_coord)

            outage_list = [
                {"start": time_to_string(start), "end": time_to_string(end)}
                for start, end in outages
            ]

            hours, minutes, total_minutes = calculate_total_time(outage_list)

            comparison_results.append({
                "queue": queue_name,
                "count": len(outage_list),
                "hours": hours,
                "minutes": minutes,
                "total_minutes": total_minutes
            })

        return comparison_results

    def _show_modern_comparison_window(self, results: List[Dict], date: str) -> None:
        """Display comparison results in modern window."""
        comparison_window = tk.Toplevel(self.root)
        comparison_window.title("Порівняння черг")
        comparison_window.geometry("700x600")
        comparison_window.configure(bg=self.colors['bg'])

        header_frame = ttk.Frame(comparison_window, style='Modern.TFrame')
        header_frame.pack(fill=tk.X, padx=30, pady=(25, 15))

        ttk.Label(header_frame,
                 text=f"Порівняння черг",
                 font=('Segoe UI', 18, 'bold'),
                 foreground=self.colors['fg'],
                 background=self.colors['bg']).pack(anchor='w')

        ttk.Label(header_frame,
                 text=f"Дата: {date}",
                 font=('Segoe UI', 12),
                 foreground=self.colors['text_secondary'],
                 background=self.colors['bg']).pack(anchor='w', pady=(5, 0))

        table_frame = tk.Frame(comparison_window, bg=self.colors['secondary_bg'],
                              highlightbackground=self.colors['border'],
                              highlightthickness=1)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 15))

        canvas = tk.Canvas(table_frame, bg=self.colors['secondary_bg'],
                          highlightthickness=0)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Card.TFrame')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        header_labels = [
            ("Черга", 150),
            ("Кількість", 100),
            ("Загальний час", 150),
            ("Хвилин", 100)
        ]

        header_row = ttk.Frame(scrollable_frame, style='Card.TFrame')
        header_row.pack(fill=tk.X, padx=15, pady=(15, 5))

        for i, (text, width) in enumerate(header_labels):
            label = tk.Label(header_row, text=text,
                           font=('Segoe UI', 10, 'bold'),
                           bg=self.colors['secondary_bg'],
                           fg=self.colors['fg'],
                           width=width//7, anchor='w')
            label.pack(side=tk.LEFT, padx=5)

        for result in results:
            row_frame = tk.Frame(scrollable_frame, bg=self.colors['input_bg'],
                               highlightbackground=self.colors['border'],
                               highlightthickness=1)
            row_frame.pack(fill=tk.X, padx=15, pady=2)

            queue_label = tk.Label(row_frame, text=result['queue'],
                                  font=('Segoe UI', 10),
                                  bg=self.colors['input_bg'],
                                  fg=self.colors['input_fg'],
                                  width=20, anchor='w')
            queue_label.pack(side=tk.LEFT, padx=5, pady=8)

            count_label = tk.Label(row_frame, text=str(result['count']),
                                  font=('Segoe UI', 10),
                                  bg=self.colors['input_bg'],
                                  fg=self.colors['accent'],
                                  width=13, anchor='w')
            count_label.pack(side=tk.LEFT, padx=5)

            time_str = f"{result['hours']}г {result['minutes']}хв"
            time_label = tk.Label(row_frame, text=time_str,
                                 font=('Segoe UI', 10),
                                 bg=self.colors['input_bg'],
                                 fg=self.colors['input_fg'],
                                 width=20, anchor='w')
            time_label.pack(side=tk.LEFT, padx=5)

            minutes_label = tk.Label(row_frame, text=f"{result['total_minutes']} хв",
                                   font=('Segoe UI', 9),
                                   bg=self.colors['input_bg'],
                                   fg=self.colors['text_secondary'],
                                   width=13, anchor='w')
            minutes_label.pack(side=tk.LEFT, padx=5)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        button_frame = ttk.Frame(comparison_window, style='Modern.TFrame')
        button_frame.pack(fill=tk.X, padx=30, pady=(0, 25))

        def copy_comparison():
            table_text = f"Порівняння черг - {date}\n\n"
            for result in results:
                table_text += f"{result['queue']:<15} {result['count']:<10} {result['hours']}г {result['minutes']}хв\n"
            comparison_window.clipboard_clear()
            comparison_window.clipboard_append(table_text)
            messagebox.showinfo("Успіх", "Таблицю скопійовано!")

        ttk.Button(button_frame, text="Копіювати",
                  command=copy_comparison, style='Secondary.TButton').pack(
                      side=tk.LEFT, padx=(0, 10))

        ttk.Button(button_frame, text="Закрити",
                  command=comparison_window.destroy, style='Secondary.TButton').pack(
                      side=tk.LEFT)

    def copy_result(self) -> None:
        """Copy analysis results to clipboard."""
        result = self.result_text.get('1.0', tk.END).strip()
        if result:
            self.root.clipboard_clear()
            self.root.clipboard_append(result)
            self.status_var.set("Результат скопійовано")
            messagebox.showinfo("Успіх", "Результат скопійовано в буфер обміну!")
        else:
            messagebox.showwarning("Попередження", "Немає результатів для копіювання!")

    def save_result(self) -> None:
        """Save analysis results to JSON file."""
        result = self.result_text.get('1.0', tk.END).strip()
        if not result:
            messagebox.showwarning("Попередження", "Немає результатів для збереження!")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON файли", "*.json"), ("Всі файли", "*.*")],
            initialfile=f"analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )

        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(result)
                self.status_var.set(f"Збережено: {os.path.basename(filename)}")
                messagebox.showinfo("Успіх", f"Результат збережено!\n\n{filename}")
            except Exception as e:
                messagebox.showerror("Помилка", f"Помилка збереження:\n{str(e)}")

    def clear_result(self) -> None:
        """Clear the result display."""
        self.result_text.delete('1.0', tk.END)
        self.status_var.set("Результати очищено")

    def export_calendar(self) -> None:
        """Export analysis results to iCalendar format."""
        result = self.result_text.get('1.0', tk.END).strip()
        if not result:
            messagebox.showwarning("Попередження", "Немає результатів для експорту!")
            return

        try:
            data = json.loads(result)
            date_str = data.get('date', '')
            queue_name = data.get('queue', '')
            outages = data.get('outages', [])

            if not outages:
                messagebox.showinfo("Інформація",
                                  "Немає відключень для експорту в календар")
                return

            date_obj = datetime.strptime(date_str, '%d.%m.%Y')

            filename = filedialog.asksaveasfilename(
                defaultextension=".ics",
                filetypes=[("Календар", "*.ics"), ("Всі файли", "*.*")],
                initialfile=f"outages_{queue_name.replace(' ', '_')}_{date_obj.strftime('%Y%m%d')}.ics"
            )

            if filename:
                self.status_var.set("Експортую в календар...")
                self.root.update_idletasks()

                success, message = CalendarExporter.export_to_ics(data, date_obj, filename)

                if success:
                    self.status_var.set(f"Календар експортовано: {len(outages)} подій")
                    try:
                        os.startfile(filename)
                        messagebox.showinfo("Успіх",
                                       f"Календар створено та відкрито!\n\n"
                                       f"Подій: {len(outages)}\n"
                                       f"Файл: {os.path.basename(filename)}")
                    except Exception:
                        messagebox.showinfo("Успіх",
                                          f"Календар створено!\n\n"
                                          f"Подій: {len(outages)}\n"
                                          f"Файл: {filename}")
                else:
                    messagebox.showerror("Помилка", f"Помилка експорту:\n{message}")
                    self.status_var.set("Помилка експорту")

        except json.JSONDecodeError:
            messagebox.showerror("Помилка", "Невірний формат JSON!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка експорту:\n{str(e)}")

