import contextlib
import os

import customtkinter as ctk
import tkinter as tk

import config
from moodle.database import Database
from gui.database_viewer import TableViewer

class MoodleDatabaseWindow(ctk.CTkToplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # Window setup
        self.title("JARS Moodle Database Explorer")
        self.resizable(True, True)
        master.eval(f"tk::PlaceWindow {self} center")

        # Frame setup
        self.inmanage_frame = MoodleDatabaseFrame(master = self, root = master)
        self.inmanage_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)

        # Hide master window
        self.master.withdraw()
        self.bind("<Destroy>", self.__on_destroy)

    def __on_destroy(self, event):
        """Shows the master window when this window is destroyed."""
        if event.widget == self:
            self.master.deiconify()
            
class MoodleDatabaseFrame(ctk.CTkFrame):
    def __init__(self, master, root, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

        # Define variables
        self.use_params = False
        self.selected_params = []
        self.param_entries = []
        self.param_labels = []

        # Categories
        categories = []
        for category in os.listdir(f"{config.get_config('moodle_custom_queries')}/"):
            if os.path.isdir(f"{config.get_config('moodle_custom_queries')}/{category}"):
                categories.append(category)

        # Dictionary of reports
        queries = []
        for filename in os.listdir(f"{config.get_config('moodle_custom_queries')}/{categories[0]}"):
            if filename.endswith(".sql"):
                queries.append(filename[:-4])

        # Define widgets
        # Text widgets
        self.lbl_title = ctk.CTkLabel(self, text = "Moodle Database Explorer", font = ("Arial", 20, "bold"))

        # Category selector
        self.category = tk.StringVar()
        self.lbl_category = ctk.CTkLabel(self, text = "Category:")
        self.cmb_category = ctk.CTkComboBox(self, values = categories, width = 250, variable = self.category)
        self.cmb_category.set(categories[0])

        # Report selector
        self.query = tk.StringVar()
        self.lbl_report = ctk.CTkLabel(self, text = "Report:")
        self.cmb_report = ctk.CTkComboBox(self, values = queries, width = 250, variable = self.query)
        self.cmb_report.set(queries[0])

        # Buttons
        self.btn_execute = ctk.CTkButton(self, text = "Execute Query", command = self.__execute_query)

        # Layout
        self.lbl_title.grid(row = 0, column = 0, columnspan = 2, padx = 10, pady = (5, 0))

        self.lbl_category.grid(row = 1, column = 0, sticky = tk.W, pady = 2)
        self.cmb_category.grid(row = 1, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_report.grid(row = 2, column = 0, sticky = tk.W, pady = 2)
        self.cmb_report.grid(row = 2, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.btn_execute.grid(row = 3, column = 1, sticky = tk.E, padx = 2, pady = 10)

        # Bindings
        self.category.trace_add("write", self.__on_category_selected)
        self.query.trace_add("write", self.__on_report_selected)

    def __on_category_selected(self, var, index, mode):
        """Updates the report selector based on the selected category."""
        queries = []
        for filename in os.listdir(f"{config.get_config('moodle_custom_queries')}/{self.category.get()}"):
            if filename.endswith(".sql"):
                queries.append(filename[:-4])
        
        self.cmb_report.configure(values = queries)
        with contextlib.suppress(IndexError):
            self.cmb_report.set(queries[0])
        tk.Misc.update_idletasks(self)

    def __on_report_selected(self, var, index, mode):
        """Opens the selected report."""
        query_file = f"{config.get_config('moodle_custom_queries')}/{self.category.get()}/{self.query.get()}.sql"

        for entry in self.param_entries:
            entry.destroy()
        
        for label in self.param_labels:
            label.destroy()

        with contextlib.suppress(FileNotFoundError):
            with open(query_file, "r") as f:
                query = f.read()
                query = query.replace("\n", " ").replace("  ", " ")
                print(f"QUERY:\n{query}\n")
            
            if query == "": 
                return
            
            if "{" in query:
                self.use_params = True
                params = []
                for param in query.split("{")[1:]:
                    params.append(param.split("}")[0])
                print(f"Parameters: {params}")

                self.selected_params.clear()
                self.selected_params = params

                self.param_entries.clear()
                self.param_labels.clear()

                for param in params:
                    self.param_labels.append(ctk.CTkLabel(self, text = f"{param.replace('_', ' ')}:"))
                    self.param_entries.append(ctk.CTkEntry(self, width = 250))

                for i in range(len(self.param_labels)):
                    self.param_labels[i].grid(row = 3 + i, column = 0, sticky = tk.W, pady = 2)
                    self.param_entries[i].grid(row = 3 + i, column = 1, sticky = tk.W, padx =  5, pady = 2)

                self.btn_execute.grid(row = 3 + len(params), column = 1, sticky = tk.E, padx = 2, pady = 10)
                tk.Misc.update_idletasks(self)
            else:
                self.use_params = False
                self.btn_execute.grid(row = 3, column = 1, sticky = tk.E, padx = 2, pady = 10)
                tk.Misc.update_idletasks(self)

    def __execute_query(self):
        print("Executing query…")

        param_dict = {}
        if self.use_params:
            for i in range(len(self.param_entries)):
                param_dict[self.selected_params[i]] = self.param_entries[i].get()

        print(f"Collected Params: {param_dict}")

        moodle = Database()
        raw_data = moodle.query_from_file(f"{config.get_config('moodle_custom_queries')}/{self.category.get()}/{self.query.get()}.sql", **param_dict)

        TableViewer(data = raw_data, report_title = self.query.get())