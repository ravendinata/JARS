import processor
import os

import customtkinter as ctk
import tkinter as tk

class OutputDialog(ctk.CTkToplevel):
    def __init__(self, master, title, file_path, content,**kwargs):
        super().__init__(master, **kwargs)

        self.file_path = file_path

        # Window Setup
        self.title(title)
        self.resizable(False, False)
        self.attributes("-topmost", True)

        ctk.CTkLabel(self, text = content).grid(row = 0, column = 0, columnspan = 3, padx = 10, pady = 10)
        ctk.CTkButton(self, text = "Open folder...", command = self.__open_folder).grid(row = 1, column = 0, padx = (10, 2), pady = 10, sticky = tk.E)
        ctk.CTkButton(self, text = "Open file...", command = self.__open_file).grid(row = 1, column = 1, padx = 2, pady = 10, sticky = tk.EW)
        ctk.CTkButton(self, text = "OK", command = self.destroy).grid(row = 1, column = 2, padx = (2, 10), pady = 10, sticky = tk.W)

    def __open_file(self):
        os.startfile(self.file_path)
        self.destroy()

    def __open_folder(self):
        os.startfile(os.path.dirname(self.file_path))
        self.destroy()

class ProcessorFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")

        self.columnconfigure(0, weight = 1)
        self.columnconfigure(1, weight = 2)
        self.columnconfigure(2, weight = 1)

        # Define widgets
        # Sourcefile path
        self.lbl_source = ctk.CTkLabel(self, text = "Source file path:")
        self.txt_source_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_source = ctk.CTkButton(self, text = "Browse...", width = 100, command = self.__browse_file)

        # Mode
        self.mode_var = tk.IntVar()
        self.lbl_mode = ctk.CTkLabel(self, text="Report file type:")
        self.rdo_mode_1 = ctk.CTkRadioButton(self, text = "Cohort All Assignment Grades", variable = self.mode_var, value = 1)
        self.rdo_mode_2 = ctk.CTkRadioButton(self, text = "Cohort All Final Grades", variable = self.mode_var, value = 2)

        # Output file path
        self.lbl_output = ctk.CTkLabel(self, text = "Output file path:")
        self.txt_output_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_output = ctk.CTkButton(self, text = "Browse...", width = 100, command = self.__save_file)

        # Adjust cell widths checkbox
        self.adjust_cell_widths_var = tk.BooleanVar()
        self.chk_adjust_cell_widths = ctk.CTkCheckBox(self, text = "Auto adjust cell widths", variable = self.adjust_cell_widths_var, onvalue = True, offvalue = False)
        self.chk_adjust_cell_widths.select() # Turn on by default

        # Process button
        self.btn_process = ctk.CTkButton(self, text = "Process", width = 100, command = self.__process)

        # Layout
        self.lbl_source.grid(row = 0, column = 0, sticky = tk.W, pady = 2)
        self.txt_source_path.grid(row = 0, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_source.grid(row = 0, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_mode.grid(row = 1, column = 0, sticky = tk.W, pady = 2)
        self.rdo_mode_1.grid(row = 2, column = 0, columnspan = 2, sticky = tk.W, padx = 5, pady = 2)
        self.rdo_mode_2.grid(row = 3, column = 0, columnspan = 2, sticky = tk.W, padx = 5, pady = 2)

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
        # Get values
        source_file = self.txt_source_path.get()
        mode = self.mode_var.get()
        output_file_path = self.txt_output_path.get()

        # Process
        proc = processor.Processor(source_file, output_file_path, mode)
        
        try:
            proc.generate_xlsx(self.adjust_cell_widths_var.get())
            OutputDialog(self.master, "Processing completed!", output_file_path, f"Done. Output file saved at {output_file_path}")
        except Exception as e:
            tk.messagebox.showerror("Error", str(e))

class Window(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Root window setup
        self.title("JARS Report Processor")

        self.processor_frame = ProcessorFrame(master = self)
        self.processor_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)

window = Window()
window.mainloop()