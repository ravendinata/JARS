"""
This program is a GUI variant of the JARS program.

It is meant to be used as a standalone application being more user-friendly 
and easier to use than the console application.
"""

__version__ = "1.2.7"
__author__ = "Raven Limadinata"

import customtkinter as ctk
import pyfiglet
import tkinter as tk
import tktooltip as tktip
from ctypes import windll
from PIL import Image
from termcolor import colored

import config
import moodle.database as moodle
from components.utility import check_word_status
from gui.report_formatter import ReportFormatterWindow
from gui.report_generator import ReportGeneratorWindow
from gui.inmanage_verifier import InManageVerifierWindow
from gui.moodle_database import MoodleDatabaseWindow
from gui.configurator import ConfiguratorWindow

global office_version
global moodle_reachable

class LauncherFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")

        # Variables
        img_attention = Image.open("resources/icons/attention.png")
        icon_attention = ctk.CTkImage(light_image = img_attention, 
                                      dark_image = img_attention, 
                                      size = (20, 20))

        # Define widgets
        # Text widgets
        self.lbl_title = ctk.CTkLabel(self, text = "JARS Report Processor", font = ("Arial", 20, "bold"))
        self.lbl_subtitle = ctk.CTkLabel(self, text = f"JAC Academic Reporting System | Version {__version__}", font = ("Arial", 10))
        self.lbl_prompt = ctk.CTkLabel(self, text = "Select a tool to open:")
        
        # Buttons
        self.btn_report_formatter = ctk.CTkButton(self, text = "Report Formatter", height = 30, command = self.__open_report_formatter)
        self.btn_report_generator = ctk.CTkButton(self, text = "Report Generator", height = 30, command = self.__open_report_generator)
        self.btn_inmanage_verifier = ctk.CTkButton(self, text = "InManage Verifier", height = 30, command = self.__open_inmanage_verifier, fg_color = "purple")
        
        self.btn_moodle_database = ctk.CTkButton(self, text = "Moodle Database Explorer", height = 30, command = self.__open_moodle_database, fg_color = "darkorange1")
        if not moodle_reachable:
            self.btn_moodle_database.configure(image = icon_attention)
            tktip.ToolTip(self.btn_moodle_database, "Cannot reach Moodle Database. Hence, function is disabled. Details are in console. Please check your connection settings or contact administrator.")
        
        self.btn_configurator = ctk.CTkButton(self, text = "Settings…", height = 30, command = self.__open_configurator, fg_color = "purple")

        # Exit button
        self.btn_exit = ctk.CTkButton(self, text = "Exit", command = self.master.destroy, fg_color = "grey")

        # Layout
        self.lbl_title.grid(row = 0, column = 0, columnspan = 2, padx = 10, pady = (5, 0))
        self.lbl_subtitle.grid(row = 1, column = 0, columnspan = 2, padx = 10, pady = (0, 10))
        self.lbl_prompt.grid(row = 2, column = 0, columnspan = 2, padx = 10, pady = 5)
        
        self.btn_report_formatter.grid(row = 3, column = 0, padx = 2, pady = (5, 2))
        self.btn_report_generator.grid(row = 3, column = 1, padx = 2, pady = (5, 2))
        self.btn_inmanage_verifier.grid(row = 4, column = 0, columnspan = 2, padx = 2, pady = 2)
        self.btn_moodle_database.grid(row = 5, column = 0, columnspan = 2, padx = 2, pady = (2, 5))

        self.btn_configurator.grid(row = 6, column = 0, columnspan = 2, padx = 2, pady = 5)
        
        self.btn_exit.grid(row = 7, column = 0, columnspan = 2, padx = 2, pady = (10, 5))

    def __open_report_formatter(self):
        ReportFormatterWindow(master = self.master)

    def __open_report_generator(self):
        next = ReportGeneratorWindow(master = self.master, office_version = office_version)
        if config.get_config("genai_api_key") == "" or config.get_config("genai_api_key") is None:
            if tk.messagebox.askyesno("API Key Missing", "You have not set an API key for the AI comment generator. Do you want to open the settings page to set the key now?"):
                ConfiguratorWindow(master = next.master)

    def __open_inmanage_verifier(self):
        InManageVerifierWindow(master = self.master)

    def __open_moodle_database(self):
        if moodle.test_connection():
            MoodleDatabaseWindow(master = self.master)
        else:
            tk.messagebox.showerror("Connection Error", "Connection to Moodle database failed. Please check your connection settings or contact administrator.")

    def __open_configurator(self):
        ConfiguratorWindow(master = self.master, caller = "main")

class Window(ctk.CTk): 
    def __init__(self):
        super().__init__()

        # Root window setup
        self.title("JARS Report Processor")
        self.eval("tk::PlaceWindow . center")

        self.processor_frame = LauncherFrame(master = self)
        self.processor_frame.grid(padx = 10, pady = 10)

"""
CHECKS
"""
print(colored(pyfiglet.figlet_format("JARS CReP", font = "slant"), "magenta"))
print(f"JARS Report Processor v{__version__}\n")

print("> Running system checks…")

# Check if MS Word is running
print("> Checking Microsoft Word status…")
is_running, office_version, error = check_word_status()
if error:
    print(f"  An error occurred: {error}")
    print("  Microsoft Office not detected.")

# Check if Moodle database is reachable
if not config.get_config("skip_moodle_check"):
    try:
        if moodle.test_connection(supress = True):
            moodle_reachable = True
            print("  Moodle database connection successful.")
        else:
            moodle_reachable = False
            print("  Cannot reach Moodle database.")
    except Exception as e:
        moodle_reachable = False
        print(f"  Moodle database connection failed due to excecption: {e}")
else:
    moodle_reachable = False
    print("  Skipping Moodle database connection check.")

print("> System checks completed.")

# Apply GUI settings
print("> Applying GUI settings…")
ctk.set_appearance_mode(config.get_config("appearance_mode"))
ctk.set_default_color_theme(config.get_config("color_theme"))

windll.shcore.SetProcessDpiAwareness(1)
window = Window()
window.mainloop()