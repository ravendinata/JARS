from typing import Tuple
import customtkinter as ctk
import tkinter as tk
import tktooltip as tktip

import config

class ConfiguratorWindow(ctk.CTkToplevel):
    def __init__(self, master, caller = None, **kwargs):
        super().__init__(master, **kwargs)

        # Window Setup
        self.title("Configurator")
        self.resizable(False, False)

        # Frame Setup
        self.config_frame = ConfiguratorFrame(master = self, root = master)
        self.config_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)

        # Hide the main window
        if caller == "main":
            master.eval(f"tk::PlaceWindow {self} center")
            self.master.withdraw()
            self.bind("<Destroy>", self.__on_close)

    def __on_close(self, event):
        """Shows the master window when this window is destroyed."""
        if event.widget == self:
            self.master.deiconify()

class ConfiguratorFrame(ctk.CTkFrame):
    def __init__(self, master, root, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

        """
        WIDGETS SETUP
        """        
        # Google GenAI section
        self.lbl_header_genai = ctk.CTkLabel(self, text = "Google GenAI settings:", font = ("Arial", 14, "bold"))

        # API Key setting
        self.lbl_api_key = ctk.CTkLabel(self, text = "API Key:")
        self.txt_api_key = ctk.CTkEntry(self, width = 300)
        self.txt_api_key.insert(0, config.get_config("genai_api_key"))

        # Buttons
        self.btn_save = ctk.CTkButton(self, text = "Save", command = self.__save_config)
        
        """
        GUI LAYOUTING
        """
        self.lbl_header_genai.grid(row = 0, column = 0, columnspan = 2, sticky = tk.W, padx = 10, pady = 5)

        self.lbl_api_key.grid(row = 1, column = 0, sticky = tk.E, padx = (10, 2), pady = 5)
        self.txt_api_key.grid(row = 1, column = 1, sticky = tk.EW, padx = (2, 10), pady = 5)

        self.btn_save.grid(row = 2, column = 1, sticky = tk.E, padx = 10, pady = 5)

    def __save_config(self):
        """Saves the configuration to the config file."""
        try:
            config.set_config("genai_api_key", self.txt_api_key.get())
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred while saving the configuration: {e}")
            return

        # Show success message
        tk.messagebox.showinfo("Success", "Configuration saved successfully.")

        self.master.destroy()