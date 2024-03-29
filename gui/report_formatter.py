import pathlib

import customtkinter as ctk
import tkinter as tk

import config
import components.report_formatter.formatter as processor
from gui.dialog import OutputDialog

class ReportFormatterWindow(ctk.CTkToplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master

        # Window Setup
        self.title("Report Formatter")
        self.resizable(False, False)

        # Frame Setup
        self.processor_frame = ReportFormatterFrame(master = self, root = master)
        self.processor_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)
        master.eval(f"tk::PlaceWindow {self} center")

        # Hide master window
        self.master.withdraw()
        self.bind("<Destroy>", self.__on_destroy)

    def __on_destroy(self, event):
        """Shows the master window when this window is destroyed."""
        if event.widget == self:
            self.master.deiconify()

class ReportFormatterFrame(ctk.CTkFrame):
    """
    A custom frame for the report processor application.

    This frame contains various widgets for selecting source file, report file type,
    output file path, and other options. It also provides functions for browsing files,
    saving files, and processing the report.

    Attributes:
        Buttons:
            btn_browse_source (CTkButton): The button for browsing the source file.
            btn_browse_output (CTkButton): The button for browsing the output folder.
            btn_process (CTkButton): The button for starting the report processing.

        Checkboxes:
            chk_adjust_cell_widths (CTkCheckBox): The checkbox for adjusting cell widths.

        Comboboxes:
            cmb_mode (CTkComboBox): The combobox for selecting the report file type.

        Labels:
            lbl_source (CTkLabel): The label for the source file path.
            lbl_mode (CTkLabel): The label for the report file type.
            lbl_output (CTkLabel): The label for the output file path.

        Text Entries:
            txt_source_path (CTkEntry): The text entry for the source file path.
            txt_output_path (CTkEntry): The text entry for the output file path.

    Methods:
        __init__(self, master, **kwargs): Initializes the ProcessorFrame.
        __browse_file(self): Opens a file dialog for browsing the source file.
        __save_file(self): Opens a file dialog for saving the output file.
        __process(self): Processes the report based on the selected options.
        __get_presets(self): Gets the list of available report file types.
    """

    def __init__(self, master, root, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

        presets = self.__get_presets()

        self.columnconfigure(0, weight = 1)
        self.columnconfigure(1, weight = 2)
        self.columnconfigure(2, weight = 1)

        # Define widgets
        # Sourcefile path
        self.lbl_source = ctk.CTkLabel(self, text = "Source file path:")
        self.txt_source_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_source = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__browse_file)

        # Mode
        self.lbl_mode = ctk.CTkLabel(self, text="Report file type:")
        self.cmb_mode = ctk.CTkComboBox(self, values = presets, width = 250)

        # Output file path
        self.lbl_output = ctk.CTkLabel(self, text = "Output file path:")
        self.txt_output_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_output = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__save_file)

        # Adjust cell widths checkbox
        self.adjust_cell_widths_var = tk.BooleanVar()
        self.chk_adjust_cell_widths = ctk.CTkCheckBox(self, text = "Auto adjust cell widths", variable = self.adjust_cell_widths_var, onvalue = True, offvalue = False)
        self.chk_adjust_cell_widths.select()  # Turn on by default

        # Process button
        self.btn_process = ctk.CTkButton(self, text = "Process", width = 100, command = self.__process)

        # Layout
        self.lbl_source.grid(row = 0, column = 0, sticky = tk.W, pady = 2)
        self.txt_source_path.grid(row = 0, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_source.grid(row = 0, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_mode.grid(row = 1, column = 0, sticky = tk.W, pady = 2)
        self.cmb_mode.grid(row = 1, column = 1, columnspan = 2, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_output.grid(row = 4, column = 0, sticky = tk.W, pady = 2)
        self.txt_output_path.grid(row = 4, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_output.grid(row = 4, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.chk_adjust_cell_widths.grid(row = 5, column = 0, columnspan = 2, sticky = tk.W, padx = 5, pady = 10)
        self.btn_process.grid(row = 5, column = 2, sticky = tk.EW, padx = 2, pady = 10)

    # UI functions
    def __browse_file(self):
        file_path = tk.filedialog.askopenfilename(defaultextension = ".csv", filetypes = [("Comma-Separated Value", "*.csv")])
        self.txt_source_path.delete(0, tk.END)  # Remove current file path in entry
        self.txt_source_path.insert(0, file_path)  # Insert the file path into entry

    def __save_file(self):
        file_path = tk.filedialog.asksaveasfilename(defaultextension = ".xlsx", filetypes = [("Microsoft Excel Document", "*.xlsx")])
        self.txt_output_path.delete(0, tk.END)  # Remove current file path in entry
        self.txt_output_path.insert(0, file_path)  # Insert the file path into entry

    def __process(self):
        """
        Processes the report based on the selected options.
        Runs the Report Formatting processor with the selected options. If an error is encountered, a messagebox will be shown.
        """
        # Get values
        source_file = self.txt_source_path.get()
        mode = self.cmb_mode.get()
        output_file_path = self.txt_output_path.get()

        # Process
        proc = processor.Formatter(source_file, output_file_path, mode)
        
        try:
            proc.generate_xlsx(self.adjust_cell_widths_var.get())
            OutputDialog(master = self.root, title = "Processing completed!", file_path = output_file_path, content = f"Done. Output file saved at {output_file_path}")
        except Exception as e:
            print(e)
            tk.messagebox.showerror("Error Encountered!", f"Please check that you have selected the correct report type!\nError: {str(e)}")

    # Helper functions
    def __get_presets(self):
        """
        Gets the list of available report file types.
        The method gets all files in the report_parser_presets folder and returns their names without the extension.
        """
        presets = []
        for preset in pathlib.Path(config.get_config("report_parser_presets")).glob("*.json"):
            presets.append(preset.stem)
        return presets