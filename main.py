import tkinter as tk
from src.ui.main_window import ScheduleAnalyzerUI


def main():
    root = tk.Tk()
    app = ScheduleAnalyzerUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
