import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image
import json
from datetime import datetime, timedelta
from analyze_schedule import analyze_row, time_to_string, resize_image
from icalendar import Calendar, Event
import pytz


class ScheduleAnalyzerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Аналізатор Графіка Відключень")
        self.root.geometry("800x700")
        self.root.minsize(700, 600)

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

    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=3, pady=(0, 25), sticky=(tk.W, tk.E))

        title_label = ttk.Label(title_frame, text="Аналізатор Графіка Відключень",
                                font=('Segoe UI', 18, 'bold'))
        title_label.pack()

        subtitle_label = ttk.Label(title_frame, text="Автоматичний аналіз та експорт в календар",
                                   font=('Segoe UI', 9), foreground='gray')
        subtitle_label.pack()

        file_frame = ttk.LabelFrame(main_frame, text="Файл графіку", padding="15")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(0, weight=1)

        file_entry = ttk.Entry(file_frame, textvariable=self.image_path, state='readonly', font=('Segoe UI', 9))
        file_entry.grid(row=0, column=0, padx=(0, 10), sticky=(tk.W, tk.E))

        file_btn = ttk.Button(file_frame, text="Вибрати файл", command=self.select_file)
        file_btn.grid(row=0, column=1)

        date_frame = ttk.LabelFrame(main_frame, text="Дата відключення", padding="15")
        date_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))

        ttk.Label(date_frame, text="Дата (РРРР-ММ-ДД):", font=('Segoe UI', 9)).grid(row=0, column=0, padx=(0, 10))
        date_entry = ttk.Entry(date_frame, textvariable=self.date_var, width=20, font=('Segoe UI', 9))
        date_entry.grid(row=0, column=1, sticky=tk.W)

        ttk.Label(date_frame, text="Наприклад: 2026-01-12", font=('Segoe UI', 8), foreground='gray').grid(
            row=0, column=2, padx=(10, 0))

        queue_frame = ttk.LabelFrame(main_frame, text="[Q] Черга відключення", padding="15")
        queue_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))

        ttk.Label(queue_frame, text="Оберіть чергу:", font=('Segoe UI', 9)).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        queue_combo = ttk.Combobox(queue_frame, textvariable=self.selected_queue,
                                    values=list(self.available_queues.keys()),
                                    state='readonly', width=35, font=('Segoe UI', 9))
        queue_combo.grid(row=0, column=1, sticky=tk.W)
        queue_combo.set("Черга 1-1")

        analyze_btn = ttk.Button(main_frame, text="Аналізувати графік",
                                command=self.analyze_schedule)
        analyze_btn.grid(row=4, column=0, columnspan=3, pady=(0, 20), ipadx=30, ipady=10)

        result_frame = ttk.LabelFrame(main_frame, text="Результат аналізу (JSON)", padding="15")
        result_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        self.result_text = scrolledtext.ScrolledText(result_frame, width=70, height=12,
                                                      wrap=tk.WORD, font=('Consolas', 9),
                                                      relief=tk.FLAT, borderwidth=2)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(0, 10))

        ttk.Button(button_frame, text="Копіювати JSON",
                  command=self.copy_result).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Зберегти JSON",
                  command=self.save_result).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Експорт в Календар (ICS)",
                  command=self.export_calendar).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Очистити",
                  command=self.clear_result).grid(row=0, column=3, padx=5)

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

            output = {
                "date": formatted_date,
                "queue": queue_name,
                "outages": results[queue_name]
            }

            json_output = json.dumps(output, ensure_ascii=False, indent=2)
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', json_output)

            messagebox.showinfo("Успіх", f"Аналіз завершено!\n\nВідключень знайдено: {len(results[queue_name])}")

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

    def auto_export_calendar(self, data, date_obj):
        queue_name = data.get('queue', '')
        outages = data.get('outages', [])

        if not outages:
            messagebox.showinfo("Успіх", "Аналіз завершено! Відключень не знайдено.")
            return

        try:
            cal = Calendar()
            cal.add('prodid', '-//Schedule Analyzer//UA')
            cal.add('version', '2.0')
            cal.add('x-wr-calname', f'Графік відключень - {queue_name}')
            cal.add('x-wr-timezone', 'Europe/Kiev')

            tz = pytz.timezone('Europe/Kiev')

            for outage in outages:
                start_time_str = outage['start']
                end_time_str = outage['end']

                start_hour, start_min = map(int, start_time_str.split(':'))
                end_hour, end_min = map(int, end_time_str.split(':'))

                start_dt = tz.localize(datetime.combine(
                    date_obj.date(),
                    datetime.min.time().replace(hour=start_hour, minute=start_min)
                ))

                if end_hour == 24 and end_min == 0:
                    end_dt = tz.localize(datetime.combine(
                        date_obj.date() + timedelta(days=1),
                        datetime.min.time()
                    ))
                else:
                    end_dt = tz.localize(datetime.combine(
                        date_obj.date(),
                        datetime.min.time().replace(hour=end_hour, minute=end_min)
                    ))

                event = Event()
                event.add('summary', f'[!] Відключення електроенергії - {queue_name}')
                event.add('dtstart', start_dt)
                event.add('dtend', end_dt)
                event.add('description',
                         f'Планове відключення електроенергії\n'
                         f'Черга: {queue_name}\n'
                         f'Час: {start_time_str} - {end_time_str}')
                event.add('location', 'Україна')
                event.add('status', 'CONFIRMED')

                from icalendar import Alarm

                alarm_15 = Alarm()
                alarm_15.add('action', 'DISPLAY')
                alarm_15.add('description', f'[!] Відключення через 15 хвилин - {queue_name}')
                alarm_15.add('trigger', timedelta(minutes=-15))
                event.add_component(alarm_15)

                alarm_5 = Alarm()
                alarm_5.add('action', 'DISPLAY')
                alarm_5.add('description', f'[!] Відключення через 5 хвилин - {queue_name}')
                alarm_5.add('trigger', timedelta(minutes=-5))
                event.add_component(alarm_5)

                cal.add_component(event)

            filename = filedialog.asksaveasfilename(
                defaultextension=".ics",
                filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")],
                initialfile=f"outages_{queue_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.ics",
                title="Зберегти календар"
            )

            if filename:
                with open(filename, 'wb') as f:
                    f.write(cal.to_ical())

                import os
                try:
                    os.startfile(filename)
                    messagebox.showinfo("Успіх",
                                   f"Календар збережено та відкрито!\n\n"
                                   f"Файл: {filename}\n"
                                   f"Події: {len(outages)}")
                except Exception as e:
                    messagebox.showinfo("Успіх",
                                      f"Календар збережено!\n\n"
                                      f"Файл: {filename}\n"
                                      f"Події: {len(outages)}")
            else:
                messagebox.showinfo("Інформація", "Аналіз завершено! Календар не збережено.")

        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка експорту календаря:\n{str(e)}")

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

            cal = Calendar()
            cal.add('prodid', '-//Schedule Analyzer//UA')
            cal.add('version', '2.0')
            cal.add('x-wr-calname', f'Графік відключень - {queue_name}')
            cal.add('x-wr-timezone', 'Europe/Kiev')

            tz = pytz.timezone('Europe/Kiev')

            for outage in outages:
                start_time_str = outage['start']
                end_time_str = outage['end']

                start_hour, start_min = map(int, start_time_str.split(':'))
                end_hour, end_min = map(int, end_time_str.split(':'))

                start_dt = tz.localize(datetime.combine(
                    date_obj.date(),
                    datetime.min.time().replace(hour=start_hour, minute=start_min)
                ))

                if end_hour == 24 and end_min == 0:
                    end_dt = tz.localize(datetime.combine(
                        date_obj.date() + timedelta(days=1),
                        datetime.min.time()
                    ))
                else:
                    end_dt = tz.localize(datetime.combine(
                        date_obj.date(),
                        datetime.min.time().replace(hour=end_hour, minute=end_min)
                    ))

                event = Event()
                event.add('summary', f'[!] Відключення електроенергії - {queue_name}')
                event.add('dtstart', start_dt)
                event.add('dtend', end_dt)
                event.add('description',
                         f'Планове відключення електроенергії\n'
                         f'Черга: {queue_name}\n'
                         f'Час: {start_time_str} - {end_time_str}')
                event.add('location', 'Україна')
                event.add('status', 'CONFIRMED')

                from icalendar import Alarm
                alarm_15 = Alarm()
                alarm_15.add('action', 'DISPLAY')
                alarm_15.add('description', f'[!] Відключення через 15 хвилин - {queue_name}')
                alarm_15.add('trigger', timedelta(minutes=-15))
                event.add_component(alarm_15)

                alarm_5 = Alarm()
                alarm_5.add('action', 'DISPLAY')
                alarm_5.add('description', f'[!] Відключення через 5 хвилин - {queue_name}')
                alarm_5.add('trigger', timedelta(minutes=-5))
                event.add_component(alarm_5)

                cal.add_component(event)

            filename = filedialog.asksaveasfilename(
                defaultextension=".ics",
                filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")],
                initialfile=f"outages_{queue_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.ics"
            )

            if filename:
                with open(filename, 'wb') as f:
                    f.write(cal.to_ical())

                import os
                try:
                    os.startfile(filename)
                    messagebox.showinfo("Успіх",
                                   f"Календар експортовано та відкрито!\n\n"
                                   f"Файл: {filename}\n"
                                   f"Події: {len(outages)}")
                except Exception as e:
                    messagebox.showinfo("Успіх",
                                      f"Календар експортовано!\n\n"
                                      f"Файл: {filename}\n"
                                      f"Події: {len(outages)}")

        except json.JSONDecodeError:
            messagebox.showerror("Помилка", "Невірний формат JSON результату!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка експорту календаря:\n{str(e)}")


def main():
    root = tk.Tk()
    app = ScheduleAnalyzerUI(root)

    style = ttk.Style()
    style.theme_use('clam')

    root.mainloop()


if __name__ == "__main__":
    main()

