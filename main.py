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
from win32com.client.dynamic import Dispatch

from gui.report_formatter import ReportFormatterWindow
from gui.report_generator import ReportGeneratorWindow
from gui.inmanage_verifier import InManageVerifierWindow
from gui.moodle_database import MoodleDatabaseWindow

global office_version

class LauncherFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")

        # Define widgets
        # Text widgets
        self.lbl_title = ctk.CTkLabel(self, text = "JARS Report Processor", font = ("Arial", 20, "bold"))
        self.lbl_subtitle = ctk.CTkLabel(self, text = "JAC Academic Reporting System | Version 1.0.0", font = ("Arial", 10))
        self.lbl_prompt = ctk.CTkLabel(self, text = "Select a tool to open:")
        
        # Buttons
        self.btn_report_formatter = ctk.CTkButton(self, text = "Report Formatter", height = 30, command = self.__open_report_formatter)
        self.btn_report_generator = ctk.CTkButton(self, text = "Report Generator", height = 30, command = self.__open_report_generator)
        self.btn_inmanage_verifier = ctk.CTkButton(self, text = "InManage Verifier", height = 30, command = self.__open_inmanage_verifier, fg_color = "purple")
        self.btn_moodle_database = ctk.CTkButton(self, text = "Moodle Database Explorer", height = 30, command = self.__open_moodle_database, fg_color = "orange")

        # Exit button
        self.btn_exit = ctk.CTkButton(self, text = "Exit", command = self.master.destroy, fg_color = "grey")

        # Layout
        self.lbl_title.grid(row = 0, column = 0, columnspan = 2, padx = 10, pady = (5, 0))
        self.lbl_subtitle.grid(row = 1, column = 0, columnspan = 2, padx = 10, pady = (0, 10))
        self.lbl_prompt.grid(row = 2, column = 0, columnspan = 2, padx = 10, pady = 5)
        
        self.btn_report_formatter.grid(row = 3, column = 0, padx = 2, pady = (5, 2))
        self.btn_report_generator.grid(row = 3, column = 1, padx = 2, pady = (5, 2))
        self.btn_inmanage_verifier.grid(row = 4, column = 0, columnspan = 2, padx = 2, pady = (2, 5))
        self.btn_moodle_database.grid(row = 5, column = 0, columnspan = 2, padx = 2, pady = (2, 5))
        
        self.btn_exit.grid(row = 6, column = 0, columnspan = 2, padx = 2, pady = (10, 5))

    def __open_report_formatter(self):
        ReportFormatterWindow(master = self.master)

    def __open_report_generator(self):
        ReportGeneratorWindow(master = self.master, office_version = office_version)

    def __open_inmanage_verifier(self):
        InManageVerifierWindow(master = self.master)

    def __open_moodle_database(self):
        MoodleDatabaseWindow(master = self.master)

class Window(ctk.CTk): 
    def __init__(self):
        super().__init__()

        # Root window setup
        self.title("JARS Report Processor")
        self.eval("tk::PlaceWindow . center")

        self.processor_frame = LauncherFrame(master = self)
        self.processor_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)

"""
CHECKS
"""
print("> Running system checks…")
try:
    word = Dispatch("Word.Application")
    office_version = word.Version
    word.Quit()
    print(f"  Microsoft Office {office_version} detected.")
except:
    office_version = None
    print("  Microsoft Office not detected.")

print("> System checks completed.")

windll.shcore.SetProcessDpiAwareness(1)
window = Window()
window.mainloop()