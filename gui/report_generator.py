import os
from datetime import date

import customtkinter as ctk
import tkinter as tk
import tktooltip as tktip
import tkcalendar as tkcal

import processor.semester_report as processor
import processor.grader_report as grader_report
import processor.helper.comment_generator_test as cgen_test
import processor.helper.language_tool_master as ltm
from gui.dialog import OutputDialog
from processor.helper.comment_generator import CommentGenerator

class ReportGeneratorWindow(ctk.CTkToplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # Window Setup
        self.title("Report Formatter")
        self.resizable(False, False)

        # Frame Setup
        self.processor_frame = ReportGeneratorFrame(master = self)
        self.processor_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)

        # Hide master window
        self.master.withdraw()
        self.bind("<Destroy>", self.__on_destroy)

    def __on_destroy(self, event):
        """Shows the master window when this window is destroyed."""
        if event.widget == self:
            self.master.deiconify()

class ReportGeneratorFrame(ctk.CTkFrame):
    """
    A custom frame for the report processor application.

    This frame contains various widgets for selecting source file, report file type,
    output file path, and other options. It also provides functions for browsing files,
    saving files, and processing the report.

    Attributes:
        Buttons:
        btn_browse_source (ctk.CTkButton): Button for browsing the source file.
        btn_browse_output (ctk.CTkButton): Button for browsing the output file path.
        btn_process (ctk.CTkButton): Button for initiating the report processing.

        Labels:
        lbl_source (ctk.CTkLabel): Label for the source file path.
        lbl_output (ctk.CTkLabel): Label for the output file path.
        lbl_generate (ctk.CTkLabel): Label for the generate option.
        lbl_student_name (ctk.CTkLabel): Label for the student name entry.
        lbl_options (ctk.CTkLabel): Label for the options section.
        lbl_progress (ctk.CTkLabel): Label for the progress tracker.
        lbl_count (ctk.CTkLabel): Label for the progress count.
        lbl_status (ctk.CTkLabel): Label for the status section.
        lbl_status_text (ctk.CTkLabel): Label for the status text.

        Text fields:
        txt_source_path (ctk.CTkEntry): Entry field for entering the source file path.
        txt_output_path (ctk.CTkEntry): Entry field for entering the output file path.
        txt_student_name (ctk.CTkEntry): Entry field for entering the student name.
        
        progress_bar (ctk.CTkProgressBar): Progress bar for tracking the progress.

        mode_var (tk.StringVar): Variable for the generate option.

    Methods:
        __init__(self, master, **kwargs): Initializes the ProcessorFrame.
        __browse_file(self): Opens a file dialog for browsing the source file.
        __save_file(self): Opens a file dialog for saving the output file.
        __opt_all_selected(self): Disables the student name entry when the generate all option is selected.
        __opt_student_selected(self): Enables the student name entry when the generate for student option is selected.
        __process(self): Processes the report based on the selected options.
        __test_source(self): Tests the source file using the comment generator test suite
        __test_paths(self, source_path, out_path): Tests if the source and output paths are valid.
        __on_progress_update(self, current, total, status_message): Updates the progress bar.
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")

        """
        WIDGETS SETUP
        """
        # Grader report file path
        self.lbl_source = ctk.CTkLabel(self, text = "Grader Report File:")
        self.txt_source_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_source = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__browse_file)

        # Output file path
        self.lbl_output = ctk.CTkLabel(self, text = "Output Folder:")
        self.txt_output_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_output = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__save_file)

        # Signature file path
        self.lbl_signature = ctk.CTkLabel(self, text = "Signature File:")
        self.txt_signature_path = ctk.CTkEntry(self, width = 250, placeholder_text = "Leave blank to not include signature")
        self.btn_browse_signature = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__browse_signature)

        # Generate all or generate for student
        self.mode_var = tk.StringVar()
        self.lbl_generate = ctk.CTkLabel(self, text = "Generate:")
        self.rdo_generate_all = ctk.CTkRadioButton(self, text = "All", variable = self.mode_var, value = "all", command = self.__opt_all_selected)
        self.rdo_generate_student = ctk.CTkRadioButton(self, text = "Student", variable = self.mode_var, value = "student", command = self.__opt_student_selected)
        self.rdo_generate_all.select()

        # Student name
        self.lbl_student_name = ctk.CTkLabel(self, text = "Student Name:")
        self.txt_student_name = ctk.CTkEntry(self, width = 250, state = tk.DISABLED)

        # Options section
        self.lbl_options = ctk.CTkLabel(self, text = "Options:")

        # Autocorrect switch button
        self.autocorrect_var = tk.IntVar()
        self.switch_autocorrect = ctk.CTkSwitch(self, text = "Autocorrect", variable = self.autocorrect_var, onvalue = 1, offvalue = 0)

        # Insert date switch button
        self.inject_date = tk.IntVar()
        self.switch_date = ctk.CTkSwitch(self, text = "Insert Date", command = self.__toggle_date_entry, variable = self.inject_date, onvalue = 1, offvalue = 0)

        # Force generation switch button
        self.force_var = tk.IntVar()
        self.switch_force = ctk.CTkSwitch(self, text = "Force Generate", variable = self.force_var, onvalue = 1, offvalue = 0)

        # Create PDF switch button
        self.create_pdf = tk.IntVar()
        self.switch_pdf = ctk.CTkSwitch(self, text = "Create PDF", variable = self.create_pdf, onvalue = 1, offvalue = 0)

        # Report date
        self.lbl_date = ctk.CTkLabel(self, text = "Report Date:")
        today = date.today()
        self.date_report = tkcal.DateEntry(self, width = 25, 
                                           background = "black", foreground = "white", 
                                           year = today.year, month = today.month, day = today.day, 
                                           font = ("Arial", 20), 
                                           date_pattern = "dd/mm/yyyy",
                                           state = tk.DISABLED)
        
        # Generate button
        self.btn_process = ctk.CTkButton(self, text = "Generate", width = 100, command = self.__process)
        self.btn_test_source = ctk.CTkButton(self, text = "Test Comment Gen", width = 100, command = self.__test_source)
        self.btn_test_source.configure(fg_color = "grey")
        self.btn_validate = ctk.CTkButton(self, text = "Validate Grader Report", width = 100, command = self.__validate)
        self.btn_validate.configure(fg_color = "grey")

        # Progress tracker
        self.lbl_progress = ctk.CTkLabel(self, text = "Progress:")
        self.progress_bar = ctk.CTkProgressBar(self, width = 250, mode = "determinate")
        self.lbl_count = ctk.CTkLabel(self, text = "0/0")
        self.progress_bar.set(0)

        # Status section
        self.lbl_status = ctk.CTkLabel(self, text = "Status:")
        self.lbl_status_text = ctk.CTkLabel(self, width = 300, justify = "left", anchor = tk.W, text = "Idle")

        # Tooltips
        tooltip_font = ("Arial", 16)
        tooltips = {
            self.btn_browse_source: "Browse for the grader report file.",
            self.btn_browse_output: "Browse for the output folder.",
            self.btn_browse_signature: "Browse for your digitized signature image file.",
            self.rdo_generate_all: "Generate reports for all students.",
            self.rdo_generate_student: "Generate report for a single student. Fill in the student name field to specify the student.",
            self.switch_autocorrect: "Enable this to automatically correct the comment grammar and spelling errors.",
            self.switch_force: "Enable this to force generate the reports and disregard grader report errors.",
            self.btn_process: "Start generating the reports.",
            self.btn_test_source: "Create a validation list of all possible comment combinations and dumps it into an Excel file",
            self.btn_validate: "Validate the grader report for errors."
        }

        for widget, message in tooltips.items():
            tktip.ToolTip(widget, message, font = tooltip_font)
        
        """
        GUI LAYOUTING
        """
        self.lbl_source.grid(row = 0, column = 0, sticky = tk.W, pady = 2)
        self.txt_source_path.grid(row = 0, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_source.grid(row = 0, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_output.grid(row = 1, column = 0, sticky = tk.W, pady = 2)
        self.txt_output_path.grid(row = 1, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_output.grid(row = 1, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_signature.grid(row = 2, column = 0, sticky = tk.W, pady = 2)
        self.txt_signature_path.grid(row = 2, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_signature.grid(row = 2, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_generate.grid(row = 3, column = 0, sticky = tk.W, pady = 2)
        self.rdo_generate_all.grid(row = 3, column = 1, sticky = tk.W, padx =  5, pady = 2)
        
        self.rdo_generate_student.grid(row = 4, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_student_name.grid(row = 5, column = 0, sticky = tk.W, pady = 2)
        self.txt_student_name.grid(row = 5, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_options.grid(row = 6, column = 0, sticky = tk.W, pady = 2)
        self.switch_autocorrect.grid(row = 6, column = 1, sticky = tk.W, padx =  5, pady = 2)
        
        self.switch_date.grid(row = 7, column = 1, sticky = tk.EW, padx = 5, pady = 2)
        
        self.switch_force.grid(row = 8, column = 1, sticky = tk.W, padx = 5, pady = 2)

        self.switch_pdf.grid(row = 9, column = 1, sticky = tk.W, padx = 5, pady = 2)

        self.lbl_date.grid(row = 10, column = 0, sticky = tk.W, pady = 2)
        self.date_report.grid(row = 10, column = 1, sticky = tk.W, padx = 5, pady = 2)

        self.lbl_progress.grid(row = 11, column = 0, sticky = tk.W, pady = 0)
        self.progress_bar.grid(row = 11, column = 1, sticky = tk.W, padx =  5, pady = 0)
        self.lbl_count.grid(row = 11, column = 2, sticky = tk.EW, padx = 2, pady = 0)

        self.lbl_status.grid(row = 12, column = 0, sticky = tk.W, pady = 0)
        self.lbl_status_text.grid(row = 12, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 0)

        self.btn_test_source.grid(row = 13, column = 0, sticky = tk.EW, padx = 2, pady = (20, 2))
        self.btn_validate.grid(row = 13, column = 1, sticky = tk.W, padx = 2, pady = (20, 2))
        self.btn_process.grid(row = 13, column = 2, sticky = tk.EW, padx = 2, pady = (20, 2))
        

    # UI functions
    def __browse_file(self):
        """Opens a file dialog for browsing the source file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select Grader Report File", defaultextension = ".xlsm", filetypes =[("Microsoft Excel Macro-Enabled Document", "*.xlsm"), ("Microsoft Excel Document", "*.xlsx")])
        self.txt_source_path.delete(0, tk.END)
        self.txt_source_path.insert(0, file_path)

    def __browse_signature(self):
        """Opens a file dialog for browsing the signature file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select Signature File", defaultextension = ".png", filetypes =[("Portable Network Graphics", "*.png"), ("Joint Photographic Experts Group", "*.jpg *.jpeg")])
        self.txt_signature_path.delete(0, tk.END)
        self.txt_signature_path.insert(0, file_path)

    def __save_file(self):
        """Opens a file dialog for saving the output file."""
        file_path = ctk.filedialog.askdirectory(title = "Select Output Folder", mustexist = True)
        self.txt_output_path.delete(0, tk.END)
        self.txt_output_path.insert(0, file_path)

    def __opt_all_selected(self):
        """Disables the student name entry when the generate all option is selected."""
        self.txt_student_name.configure(state = tk.DISABLED)

    def __opt_student_selected(self):
        """Enables the student name entry when the generate for student option is selected."""
        self.txt_student_name.configure(state = tk.NORMAL)
        self.txt_student_name.focus_set()

    def __toggle_date_entry(self):
        """Enables or disables the date entry field when the insert date option is selected or not."""
        if self.inject_date.get() == 1:
            self.date_report.configure(state = tk.NORMAL)
        else:
            self.date_report.configure(state = tk.DISABLED)

    def __process(self):
        """
        Processes the report based on the selected options.
        
        It runs the report processor based on the selected options. It will run the report processor
        for all students if the generate all option is selected. It will run the report processor for
        a single student if the generate for student option is selected.
        """
        if not self.__test_paths():
            return
        
        source_file = self.txt_source_path.get()
        output_file_path = self.txt_output_path.get()
        signature_file = self.txt_signature_path.get() if self.txt_signature_path.get() != "" else None
        date = self.date_report.get_date() if self.inject_date.get() == 1 else None

        gr = grader_report.GraderReport(source_file)
        proc = processor.Generator(output_file_path, gr, date, signature_file)

        mode = self.mode_var.get()
        autocorrect = True if self.autocorrect_var.get() == 1 else False
        force = True if self.force_var.get() == 1 else False
        pdf = True if self.create_pdf.get() == 1 else False
        
        if mode == "all":
            proc.generate_all(callback = self.__on_progress_update, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
        elif mode == "student":
            student_name = self.txt_student_name.get()
            self.__on_progress_update(0, 1, f"Generating report for {student_name}…")
            proc.generate_for_student(student_name = student_name, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
            self.__on_progress_update(1, 1, f"Done!")
            output_file_path = f"{output_file_path}/{student_name}.docx"

        self.lbl_count.configure(text = "Done!")
        self.lbl_status_text.configure(text = "Report generation completed successfully.")

        OutputDialog(master = self.master, title = "Report Generation Complete", file_path = output_file_path, content = "Report generation completed successfully.")

    def __test_source(self):
        """
        Tests the source file using the comment generator test suite.
        It runs the comment generator test suite on the source file and saves the result to an Excel file.
        """
        if not self.__test_paths():
            return
        
        source_file = self.txt_source_path.get()
        output_file_path = self.txt_output_path.get()
        file_name = source_file.split("/")[-1].split(".")[0]
        output_file_path = f"{output_file_path}/CGen_Test_{file_name}.xlsx"
        
        self.lbl_status_text.configure(text = "Running comment generator test suite. Please wait…")
        tk.Misc.update_idletasks(self)
        
        cgen_test.run(source_file_path = source_file, output_file_path = output_file_path, ltm = ltm, CommentGenerator = CommentGenerator)
        
        OutputDialog(master = self.master, title = "Comment Generator Test Complete", file_path = output_file_path, content = "Comment generator test completed successfully.")
        self.lbl_status_text.configure(text = "Comment generator test completed successfully.")
        tk.Misc.update_idletasks(self)

    def __validate(self):
        source_file = self.txt_source_path.get()
        
        if not os.path.isfile(source_file):
            tk.messagebox.showerror("Error", "Please select a valid source file.")
            return
        
        grader_report.GraderReport(source_file)
        self.lbl_count.configure(text = "Done!")
        self.lbl_status_text.configure(text = "Validation completed. Check console/terminal for details.")

    def __test_paths(self):
        """
        Tests if the source and output paths are valid.

        Returns:
            bool: True if the paths are valid, False otherwise.
        """
        if not os.path.isfile(self.txt_source_path.get()):
            tk.messagebox.showerror("Error", "Please select a valid source file.")
            return False
        
        if not os.path.isdir(self.txt_output_path.get()):
            tk.messagebox.showerror("Error", "Please select a valid output folder.")
            return False
        
        if not os.path.isfile(self.txt_signature_path.get()) and self.txt_signature_path.get() != "":
            tk.messagebox.showerror("Error", "Please select a valid signature file.")
            return False
        
        return True

    def __on_progress_update(self, current, total, status_message):
        """
        CALLBACK: Updates the progress bar and other progress indicators.      
        
        This is a callback function for the report processor. It is called when the report processor updates its progress.
        It updates the progress bar and the progress count label.

        Args:
            current (int): The current progress.
            total (int): The total progress. Or, the value in which progress is 100%.
            status_message (str): The status message to display.
        """
        progress = current / total
        
        self.progress_bar.set(progress)
        self.lbl_count.configure(text = f"{current}/{total}")

        self.lbl_status_text.configure(text = status_message)

        tk.Misc.update_idletasks(self)