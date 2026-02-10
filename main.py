"""
Schedule Analyzer Application Entry Point.

Desktop application for analyzing power outage schedules from images
with automatic time detection and calendar export functionality.
"""

import tkinter as tk
from src.ui.main_window import ScheduleAnalyzerUI


def main() -> None:
    """Initialize and run the application."""
    root = tk.Tk()
    app = ScheduleAnalyzerUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()


