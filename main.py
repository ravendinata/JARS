"""
This program is a GUI variant of the JARS program.

It is meant to be used as a standalone application.
The GUI application is only suitable for processing a single file at a time.
But it is more user-friendly and easier to use than the console application.
"""

__version__ = "1.0.0"
__author__ = "Raven Limadinata"

import customtkinter as ctk
import tkinter as tk
from ctypes import windll

from gui.report_formatter import ReportFormatterWindow
from gui.report_generator import ReportGeneratorWindow

class LauncherFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")

        # Define widgets
        # Text widgets
        self.lbl_title = ctk.CTkLabel(self, text = "JARS Report Processor", font = ("Arial", 20, "bold"))
        self.lbl_subtitle = ctk.CTkLabel(self, text = "JAC Academic Reporting System | Version 1.0.0", font = ("Arial", 10))
        self.lbl_prompt = ctk.CTkLabel(self, text = "Select a tool to open:")
        
        # Buttons
        self.btn_report_formatter = ctk.CTkButton(self, text = "Report Formatter", command = self.__open_report_formatter)
        self.btn_report_generator = ctk.CTkButton(self, text = "Report Generator", command = self.__open_report_generator)

        # Exit button
        self.btn_exit = ctk.CTkButton(self, text = "Exit", command = self.master.destroy)

        # Layout
        self.lbl_title.grid(row = 0, column = 0, columnspan = 2, padx = 10, pady = (5, 0))
        self.lbl_subtitle.grid(row = 1, column = 0, columnspan = 2, padx = 10, pady = (0, 10))
        self.lbl_prompt.grid(row = 2, column = 0, columnspan = 2, padx = 10, pady = 5)
        
        self.btn_report_formatter.grid(row = 3, column = 0, padx = 2, pady = 5)
        self.btn_report_generator.grid(row = 3, column = 1, padx = 2, pady = 5)
        
        self.btn_exit.grid(row = 4, column = 0, columnspan = 2, padx = 2, pady = 5)

    def __open_report_formatter(self):
        ReportFormatterWindow(master = self.master)

    def __open_report_generator(self):
        ReportGeneratorWindow(master = self.master)

class Window(ctk.CTk): 
    def __init__(self):
        super().__init__()

        # Root window setup
        self.title("JARS Report Processor")
        self.eval("tk::PlaceWindow . center")

        self.processor_frame = LauncherFrame(master = self)
        self.processor_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)

windll.shcore.SetProcessDpiAwareness(1)
window = Window()
window.mainloop()