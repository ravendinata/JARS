import sys
import webbrowser

import customtkinter as ctk
import tkinter as tk

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

        self.grab_set()

    def __on_close(self, event):
        """Shows the master window when this window is destroyed."""
        if event.widget == self:
            self.master.deiconify()

class ConfiguratorFrame(ctk.CTkFrame):
    require_restart = False

    def __init__(self, master, root, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

        """
        WIDGETS SETUP
        """        
        # General section
        self.lbl_header_general = ctk.CTkLabel(self, text = "General settings:", font = ("Arial", 14, "bold"))

        # Appearance mode setting
        self.lbl_appearance = ctk.CTkLabel(self, text = "Appearance:")
        self.cmb_appearance = ctk.CTkComboBox(self, values = ["System", "Light", "Dark"], width = 300, state = "readonly")
        current_mode = config.get_config("appearance_mode")
        current_mode = current_mode[0].upper() + current_mode[1:]
        self.cmb_appearance.set(current_mode)

        # Color theme setting
        self.color_theme = tk.StringVar()
        self.lbl_color_theme = ctk.CTkLabel(self, text = "Color Theme:")
        self.cmb_color_theme = ctk.CTkComboBox(self, values = ["Blue", "Green"], width = 300, state = "readonly", variable = self.color_theme)
        current_theme = config.get_config("color_theme").replace("-", " ")
        current_theme = " ".join([word[0].upper() + word[1:] for word in current_theme.split()])
        self.cmb_color_theme.set(current_theme)

        # Skip Moodle check setting
        self.skip_moodle_check = tk.BooleanVar(value = config.get_config("skip_moodle_check"))
        self.chk_skip_moodle_check = ctk.CTkCheckBox(self, variable = self.skip_moodle_check, text = "Skip Moodle database connection check")

        # Google GenAI section
        self.lbl_header_genai = ctk.CTkLabel(self, text = "Google GenAI (Gemini) settings:", font = ("Arial", 14, "bold"))

        # API Key setting
        self.lbl_api_key = ctk.CTkLabel(self, text = "API Key:")
        self.txt_api_key = ctk.CTkEntry(self, width = 300)
        self.txt_api_key.insert(0, config.get_config("genai_api_key"))
        
        # Redirect to Google AI dashboard URL
        self.btn_genai_dashboard = ctk.CTkButton(self, text = "Get API key from Google AI Dashboard", fg_color = "grey",
                                                 command = lambda: webbrowser.open("https://makersuite.google.com/app/apikey"))

        # Buttons
        self.btn_save = ctk.CTkButton(self, text = "Save", command = self.__save_config)

        # Bindings
        self.color_theme.trace_add("write", self.__on_color_theme_change)
        
        """
        GUI LAYOUTING
        """
        self.lbl_header_general.grid(row = 0, column = 0, columnspan = 2, sticky = tk.W, padx = 10, pady = 5)

        self.lbl_appearance.grid(row = 1, column = 0, sticky = tk.E, padx = (10, 2), pady = 5)
        self.cmb_appearance.grid(row = 1, column = 1, sticky = tk.EW, padx = (2, 10), pady = 5)

        self.lbl_color_theme.grid(row = 2, column = 0, sticky = tk.E, padx = (10, 2), pady = 5)
        self.cmb_color_theme.grid(row = 2, column = 1, sticky = tk.EW, padx = (2, 10), pady = 5)

        self.chk_skip_moodle_check.grid(row = 3, column = 1, sticky = tk.EW, padx = (2, 10), pady = 5)

        self.lbl_header_genai.grid(row = 4, column = 0, columnspan = 2, sticky = tk.W, padx = 10, pady = 5)

        self.lbl_api_key.grid(row = 5, column = 0, sticky = tk.E, padx = (10, 2), pady = 5)
        self.txt_api_key.grid(row = 5, column = 1, sticky = tk.EW, padx = (2, 10), pady = 5)
        
        self.btn_genai_dashboard.grid(row = 6, column = 0, columnspan = 2, sticky = tk.EW, padx = 10, pady = 5)

        self.btn_save.grid(row = 7, column = 1, sticky = tk.E, padx = 10, pady = 5)

    def __save_config(self):
        """Saves the configuration to the config file."""
        new_settings = None

        try:
            # Saving to config file
            new_settings = config.set_config("appearance_mode", self.cmb_appearance.get().lower())
            new_settings = config.set_config("color_theme", self.color_theme.get().lower().replace(" ", "-"), source = new_settings)
            new_settings = config.set_config("genai_api_key", self.txt_api_key.get(), source = new_settings)
            new_settings = config.set_config("skip_moodle_check", self.skip_moodle_check.get(), source = new_settings)

            # Extra steps
            ctk.set_appearance_mode(self.cmb_appearance.get().lower())
            ctk.set_default_color_theme(self.cmb_color_theme.get().lower().replace(" ", "-"))

            # Save the configuration
            config.save_config(new_settings)
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred while saving the configuration: {e}")
            return

        # Show success message
        if self.require_restart:
            tk.messagebox.showinfo("Success", "Configuration saved successfully.\nSome changes require the program to be restarted.\nProgram will now exit. Please manually restart the application.")
            sys.exit()
        else:
            tk.messagebox.showinfo("Success", "Configuration saved successfully.")

        self.master.destroy()

    def __on_color_theme_change(self, var, index, mode):
        """Changes the color theme based on the selected value."""
        self.require_restart = True