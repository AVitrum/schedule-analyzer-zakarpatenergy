import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image
import json
from datetime import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from analyze_schedule import analyze_row, time_to_string, resize_image
from src.themes.theme_manager import ModernTheme
from src.utils.calendar_export import CalendarExporter
from src.utils.time_calculator import calculate_total_time, queue_sort_key


class ScheduleAnalyzerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Аналізатор Графіка Відключень")
        self.root.geometry("900x750")
        self.root.minsize(750, 650)
        
        self.dark_mode = ModernTheme.detect_system_theme()
        self.colors = ModernTheme.get_theme_colors(self.dark_mode)
        
        self.setup_modern_style()
        self.center_window()

        self.image_path = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.selected_queue = tk.StringVar()

        self.available_queues = {
            "Черга 1-1": 90,
            "Черга 1-2": 109,
            "Черга 2-1": 127,
            "Черга 2-2": 146,
            "Черга 3-1": 170,
            "Черга 3-2": 190,
            "Черга 4-1": 214,
            "Черга 4-2": 233,
            "Черга 5-1": 255,
            "Черга 5-2": 275,
            "Черга 6-1": 298,
            "Черга 6-2": 315,
        }

        self.create_widgets()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_modern_style(self):
        style = ttk.Style()
        
        style.theme_use('clam')
        
        self.root.configure(bg=self.colors['bg'])
        
        style.configure('Modern.TFrame', 
                       background=self.colors['bg'])
        
        style.configure('Card.TFrame',
                       background=self.colors['secondary_bg'],
                       relief='flat',
                       borderwidth=1)
        
        style.configure('Modern.TLabel',
                       background=self.colors['bg'],
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 10))
        
        style.configure('Title.TLabel',
                       background=self.colors['bg'],
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 20, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=self.colors['bg'],
                       foreground=self.colors['text_secondary'],
                       font=('Segoe UI', 9))
        
        style.configure('Card.TLabel',
                       background=self.colors['secondary_bg'],
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 10))
        
        style.configure('CardTitle.TLabel',
                       background=self.colors['secondary_bg'],
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 11, 'bold'))
        
        style.configure('Modern.TLabelframe',
                       background=self.colors['secondary_bg'],
                       foreground=self.colors['fg'],
                       borderwidth=0,
                       relief='flat')
        
        style.configure('Modern.TLabelframe.Label',
                       background=self.colors['secondary_bg'],
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 10, 'bold'))
        
        style.configure('Modern.TEntry',
                       fieldbackground=self.colors['input_bg'],
                       foreground=self.colors['input_fg'],
                       insertcolor=self.colors['input_fg'],
                       borderwidth=1,
                       relief='solid')
        
        style.map('Modern.TEntry',
                 fieldbackground=[('readonly', self.colors['input_bg']),
                                 ('disabled', self.colors['input_bg'])],
                 foreground=[('readonly', self.colors['input_fg']),
                            ('disabled', self.colors['text_secondary'])])
        
        style.configure('Modern.TCombobox',
                       fieldbackground=self.colors['input_bg'],
                       background=self.colors['button_bg'],
                       foreground=self.colors['input_fg'],
                       selectbackground=self.colors['accent'],
                       selectforeground=self.colors['button_fg'],
                       borderwidth=1,
                       relief='solid',
                       arrowcolor=self.colors['fg'],
                       insertcolor=self.colors['input_fg'])
        
        style.map('Modern.TCombobox',
                 fieldbackground=[('readonly', self.colors['input_bg']),
                                 ('disabled', self.colors['input_bg'])],
                 foreground=[('readonly', self.colors['input_fg']),
                            ('disabled', self.colors['text_secondary'])],
                 background=[('readonly', self.colors['button_bg']),
                            ('disabled', self.colors['input_bg'])],
                 selectbackground=[('readonly', self.colors['accent'])],
                 selectforeground=[('readonly', self.colors['button_fg'])],
                 arrowcolor=[('disabled', self.colors['text_secondary'])])
        
        style.configure('Primary.TButton',
                       background=self.colors['button_bg'],
                       foreground=self.colors['button_fg'],
                       borderwidth=0,
                       focuscolor='none',
                       font=('Segoe UI', 10, 'bold'),
                       relief='flat',
                       padding=10)
        
        style.map('Primary.TButton',
                 background=[('active', self.colors['button_hover']),
                           ('pressed', self.colors['accent']),
                           ('!active', self.colors['button_bg'])],
                 foreground=[('active', self.colors['button_fg']),
                           ('pressed', self.colors['button_fg']),
                           ('disabled', self.colors['text_secondary']),
                           ('!active', self.colors['button_fg'])])
        
        style.configure('Secondary.TButton',
                       background=self.colors['secondary_bg'],
                       foreground=self.colors['fg'],
                       borderwidth=1,
                       focuscolor='none',
                       font=('Segoe UI', 9),
                       relief='solid',
                       bordercolor=self.colors['border'],
                       padding=5)
        
        style.map('Secondary.TButton',
                 background=[('active', self.colors['input_bg']),
                           ('!active', self.colors['secondary_bg'])],
                 foreground=[('active', self.colors['fg']),
                           ('pressed', self.colors['fg']),
                           ('disabled', self.colors['text_secondary']),
                           ('!active', self.colors['fg'])])

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding="30", style='Modern.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        title_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        title_frame.grid(row=0, column=0, columnspan=3, pady=(0, 30), sticky=(tk.W, tk.E))

        title_label = ttk.Label(title_frame, text="Аналізатор Графіка Відключень",
                                style='Title.TLabel')
        title_label.pack()

        subtitle_label = ttk.Label(title_frame, text="Автоматичний аналіз та експорт в календар",
                                   style='Subtitle.TLabel')
        subtitle_label.pack(pady=(5, 0))

        file_card = ttk.LabelFrame(main_frame, text="Файл графіку", 
                                   padding="20", style='Modern.TLabelframe')
        file_card.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        file_card.columnconfigure(0, weight=1)

        file_entry = ttk.Entry(file_card, textvariable=self.image_path, 
                              state='readonly', font=('Segoe UI', 10),
                              style='Modern.TEntry')
        file_entry.grid(row=0, column=0, padx=(0, 15), sticky=(tk.W, tk.E), ipady=5)

        file_btn = ttk.Button(file_card, text="Вибрати файл", 
                             command=self.select_file, style='Secondary.TButton')
        file_btn.grid(row=0, column=1, ipady=5, ipadx=15)

        date_card = ttk.LabelFrame(main_frame, text="Дата відключення", 
                                   padding="20", style='Modern.TLabelframe')
        date_card.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        date_card.columnconfigure(1, weight=1)

        ttk.Label(date_card, text="Дата (РРРР-ММ-ДД):", 
                 style='Card.TLabel').grid(row=0, column=0, padx=(0, 15), sticky=tk.W)
        
        date_entry = ttk.Entry(date_card, textvariable=self.date_var, 
                              width=20, font=('Segoe UI', 10),
                              style='Modern.TEntry')
        date_entry.grid(row=0, column=1, sticky=tk.W, ipady=5)

        ttk.Label(date_card, text="Наприклад: 2026-01-12", 
                 font=('Segoe UI', 9), 
                 foreground=self.colors['text_secondary'],
                 background=self.colors['secondary_bg']).grid(
            row=0, column=2, padx=(15, 0))

        queue_card = ttk.LabelFrame(main_frame, text="Черга відключення", 
                                    padding="20", style='Modern.TLabelframe')
        queue_card.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 25))
        queue_card.columnconfigure(1, weight=1)

        ttk.Label(queue_card, text="Оберіть чергу:", 
                 style='Card.TLabel').grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        queue_combo = ttk.Combobox(queue_card, textvariable=self.selected_queue,
                                    values=list(self.available_queues.keys()),
                                    state='readonly', width=35, font=('Segoe UI', 10),
                                    style='Modern.TCombobox')
        queue_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), ipady=5)
        queue_combo.set("Черга 1-1")

        button_row = ttk.Frame(main_frame, style='Modern.TFrame')
        button_row.grid(row=4, column=0, columnspan=3, pady=(0, 20))

        analyze_btn = ttk.Button(button_row, text="Аналізувати графік",
                                command=self.analyze_schedule, style='Primary.TButton')
        analyze_btn.grid(row=0, column=0, padx=10, ipadx=40, ipady=12)

        compare_btn = ttk.Button(button_row, text="Порівняти всі черги",
                                command=self.compare_all_queues, style='Secondary.TButton')
        compare_btn.grid(row=0, column=1, padx=10, ipadx=30, ipady=12)

        result_card = ttk.LabelFrame(main_frame, text="Результат аналізу (JSON)", 
                                     padding="20", style='Modern.TLabelframe')
        result_card.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        result_card.columnconfigure(0, weight=1)
        result_card.rowconfigure(0, weight=1)

        self.result_text = scrolledtext.ScrolledText(
            result_card, 
            width=70, 
            height=12,
            wrap=tk.WORD, 
            font=('Consolas', 10),
            bg=self.colors['input_bg'],
            fg=self.colors['input_fg'],
            relief=tk.FLAT,
            borderwidth=0,
            insertbackground=self.colors['fg']
        )
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        button_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        button_frame.grid(row=6, column=0, columnspan=3, pady=(0, 10))

        ttk.Button(button_frame, text="Копіювати JSON",
                  command=self.copy_result, style='Secondary.TButton',
                  width=18).grid(row=0, column=0, padx=5, ipady=5)
        ttk.Button(button_frame, text="Зберегти JSON",
                  command=self.save_result, style='Secondary.TButton',
                  width=18).grid(row=0, column=1, padx=5, ipady=5)
        ttk.Button(button_frame, text="Експорт в Календар",
                  command=self.export_calendar, style='Secondary.TButton',
                  width=20).grid(row=0, column=2, padx=5, ipady=5)
        ttk.Button(button_frame, text="Очистити",
                  command=self.clear_result, style='Secondary.TButton',
                  width=15).grid(row=0, column=3, padx=5, ipady=5)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Виберіть файл графіку",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.image_path.set(filename)

    def compare_all_queues(self):
        if not self.image_path.get():
            messagebox.showerror("Помилка", "Будь ласка, виберіть файл графіку!")
            return

        if not self.date_var.get():
            messagebox.showerror("Помилка", "Будь ласка, вкажіть дату!")
            return

        try:
            date_obj = datetime.strptime(self.date_var.get(), '%Y-%m-%d')
            formatted_date = date_obj.strftime('%d.%m.%Y')
        except ValueError:
            messagebox.showerror("Помилка", "Невірний формат дати!")
            return

        try:
            img = Image.open(self.image_path.get())
            img = resize_image(img)

            comparison_results = []
            
            for queue_name, y_coord in self.available_queues.items():
                outages = analyze_row(img, y_coord)
                
                outage_list = [
                    {
                        "start": time_to_string(start),
                        "end": time_to_string(end)
                    }
                    for start, end in outages
                ] if outages else []
                
                hours, minutes, total_minutes = calculate_total_time(outage_list)
                
                comparison_results.append({
                    "queue": queue_name,
                    "count": len(outage_list),
                    "hours": hours,
                    "minutes": minutes,
                    "total_minutes": total_minutes
                })
            
            comparison_results.sort(key=queue_sort_key)
            
            table_header = f"{'Черга':<15} {'Кількість':<12} {'Загальний час':<20}\n"
            table_header += "=" * 50 + "\n"
            
            table_rows = ""
            for result in comparison_results:
                queue = result['queue']
                count = result['count']
                time_str = f"{result['hours']} год {result['minutes']} хв"
                table_rows += f"{queue:<15} {count:<12} {time_str:<15}\n"
            
            table_message = f"Порівняння черг ({formatted_date})\n\n{table_header}{table_rows}"
            
            comparison_window = tk.Toplevel(self.root)
            comparison_window.title("Порівняння черг")
            comparison_window.geometry("600x500")
            comparison_window.configure(bg=self.colors['bg'])
            
            title_label = ttk.Label(comparison_window, 
                                   text=f"Порівняння черг - {formatted_date}",
                                   style='Title.TLabel',
                                   font=('Segoe UI', 14, 'bold'))
            title_label.pack(pady=(20, 10))
            
            frame = ttk.Frame(comparison_window, style='Modern.TFrame', padding=20)
            frame.pack(fill=tk.BOTH, expand=True)
            
            text_widget = scrolledtext.ScrolledText(
                frame,
                wrap=tk.WORD,
                font=('Consolas', 11),
                bg=self.colors['input_bg'],
                fg=self.colors['input_fg'],
                relief=tk.FLAT,
                borderwidth=0,
                insertbackground=self.colors['fg']
            )
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert('1.0', table_message)
            text_widget.config(state=tk.DISABLED)
            
            btn_frame = ttk.Frame(comparison_window, style='Modern.TFrame')
            btn_frame.pack(pady=(10, 20))
            
            def copy_table():
                comparison_window.clipboard_clear()
                comparison_window.clipboard_append(table_message)
                messagebox.showinfo("Успіх", "Таблицю скопійовано в буфер обміну!")
            
            ttk.Button(btn_frame, text="Копіювати таблицю",
                      command=copy_table, style='Secondary.TButton',
                      width=20).pack(side=tk.LEFT, padx=5, ipady=5)
            ttk.Button(btn_frame, text="Закрити",
                      command=comparison_window.destroy, style='Secondary.TButton',
                      width=15).pack(side=tk.LEFT, padx=5, ipady=5)

        except FileNotFoundError:
            messagebox.showerror("Помилка", "Файл не знайдено!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка під час порівняння:\n{str(e)}")

    def analyze_schedule(self):
        if not self.image_path.get():
            messagebox.showerror("Помилка", "Будь ласка, виберіть файл графіку!")
            return

        if not self.date_var.get():
            messagebox.showerror("Помилка", "Будь ласка, вкажіть дату!")
            return

        if not self.selected_queue.get():
            messagebox.showerror("Помилка", "Будь ласка, виберіть чергу для аналізу!")
            return

        try:
            date_obj = datetime.strptime(self.date_var.get(), '%Y-%m-%d')
            formatted_date = date_obj.strftime('%d.%m.%Y')
        except ValueError:
            messagebox.showerror("Помилка", "Невірний формат дати! Використовуйте РРРР-ММ-ДД (наприклад: 2026-01-12)")
            return

        try:
            img = Image.open(self.image_path.get())
            img = resize_image(img)

            queue_name = self.selected_queue.get()
            y_coord = self.available_queues[queue_name]
            outages = analyze_row(img, y_coord)

            results = {}
            if outages:
                results[queue_name] = [
                    {
                        "start": time_to_string(start),
                        "end": time_to_string(end)
                    }
                    for start, end in outages
                ]
            else:
                results[queue_name] = []

            hours, minutes, total_minutes = calculate_total_time(results[queue_name])
            
            output = {
                "date": formatted_date,
                "queue": queue_name,
                "outages": results[queue_name],
                "total_outage_time": {
                    "hours": hours,
                    "minutes": minutes,
                    "total_minutes": total_minutes,
                    "formatted": f"{hours} год {minutes} хв"
                }
            }

            json_output = json.dumps(output, ensure_ascii=False, indent=2)
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', json_output)

            messagebox.showinfo("Успіх", 
                f"Аналіз завершено!\n\n"
                f"Відключень знайдено: {len(results[queue_name])}\n"
                f"Загальний час: {hours} год {minutes} хв ({total_minutes} хв)")

        except FileNotFoundError:
            messagebox.showerror("Помилка", "Файл не знайдено!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка під час аналізу:\n{str(e)}")

    def copy_result(self):
        result = self.result_text.get('1.0', tk.END).strip()
        if result:
            self.root.clipboard_clear()
            self.root.clipboard_append(result)
            messagebox.showinfo("Успіх", "Результат скопійовано в буфер обміну!")
        else:
            messagebox.showwarning("Попередження", "Немає результатів для копіювання!")

    def save_result(self):
        result = self.result_text.get('1.0', tk.END).strip()
        if not result:
            messagebox.showwarning("Попередження", "Немає результатів для збереження!")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=f"schedule_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )

        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(result)
                messagebox.showinfo("Успіх", f"Результат збережено у файл:\n{filename}")
            except Exception as e:
                messagebox.showerror("Помилка", f"Помилка збереження файлу:\n{str(e)}")

    def clear_result(self):
        self.result_text.delete('1.0', tk.END)

    def export_calendar(self):
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
                messagebox.showinfo("Інформація", "Немає відключень для експорту в календар")
                return

            date_obj = datetime.strptime(date_str, '%d.%m.%Y')

            filename = filedialog.asksaveasfilename(
                defaultextension=".ics",
                filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")],
                initialfile=f"outages_{queue_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.ics"
            )

            if filename:
                success, message = CalendarExporter.export_to_ics(data, date_obj, filename)
                
                if success:
                    try:
                        os.startfile(filename)
                        messagebox.showinfo("Успіх",
                                       f"Календар експортовано та відкрито!\n\n"
                                       f"Файл: {filename}\n"
                                       f"Події: {len(outages)}")
                    except Exception:
                        messagebox.showinfo("Успіх",
                                          f"Календар експортовано!\n\n"
                                          f"Файл: {filename}\n"
                                          f"Події: {len(outages)}")
                else:
                    messagebox.showerror("Помилка", f"Помилка експорту:\n{message}")

        except json.JSONDecodeError:
            messagebox.showerror("Помилка", "Невірний формат JSON результату!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка експорту календаря:\n{str(e)}")
