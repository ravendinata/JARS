import os

import customtkinter as ctk
import tkinter as tk

class OutputDialog(ctk.CTkToplevel):
    """
    A custom message box for the report processor application.

    This message box is displayed after the report processing is completed.
    It displays the output file path and a message. It also provides buttons for opening the output file and the output folder.

    Attributes:
        file_path (str): The path of the output file.
        content (str): The message to display.

    Methods:
        __init__(self, master, title, file_path, content, **kwargs): Initializes the OutputDialog.
        __open_file(self): Opens the output file.
        __open_folder(self): Opens the output folder.

    TODO:
        Move this class to a separate custom widget file ?
    """

    def __init__(self, master, title, file_path, content,**kwargs):
        super().__init__(master, **kwargs)

        self.file_path = file_path

        # Window Setup
        self.title(title)
        self.resizable(False, False)
        self.attributes("-topmost", True)
        master.eval(f"tk::PlaceWindow {self} center")

        ctk.CTkLabel(self, text = content).grid(row = 0, column = 0, columnspan = 3, padx = 10, pady = 10)

        if file_path:
            ctk.CTkButton(self, text = "Open folder", command = self.__open_folder).grid(row = 1, column = 0, padx = (10, 2), pady = 10, sticky = tk.E)
        
        if os.path.isfile(file_path):
            ctk.CTkButton(self, text = "Open file", command = self.__open_file).grid(row = 1, column = 1, padx = 2, pady = 10, sticky = tk.EW)
        
        ctk.CTkButton(self, text = "OK", command = self.destroy).grid(row = 1, column = 2, padx = (2, 10), pady = 10, sticky = tk.W)

    def __open_file(self):
        """Opens the output file using the default program."""
        self.attributes("-topmost", False)
        os.startfile(self.file_path)

    def __open_folder(self):
        """Opens the output folder using the default program."""
        self.attributes("-topmost", False)
        if os.path.isfile(self.file_path):
            os.startfile(os.path.dirname(self.file_path))
        else:
            os.startfile(self.file_path)