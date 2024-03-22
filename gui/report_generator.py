import os
from datetime import date

import customtkinter as ctk
import tkinter as tk
import tktooltip as tktip
import tkcalendar as tkcal

import config
import components.report_generator.semester_report as processor
import components.common.grader_report as grader_report
import components.report_generator.comment_generator_test as cgen_test
from gui.dialog import OutputDialog
from components.report_generator.comment_generator import CommentGenerator

class ReportGeneratorWindow(ctk.CTkToplevel):
    def __init__(self, master, office_version, **kwargs):
        super().__init__(master, **kwargs)

        # Window Setup
        self.title("Report Formatter")
        self.resizable(False, False)

        # Frame Setup
        self.processor_frame = ReportGeneratorFrame(master = self, root = master, office_version = office_version)
        self.processor_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)
        master.eval(f"tk::PlaceWindow {self} center")

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
            btn_browse_source (CTkButton): The button for browsing the source file.
            btn_browse_output (CTkButton): The button for browsing the output folder.
            btn_browse_signature (CTkButton): The button for browsing the signature file.
            btn_test_source (CTkButton): The button for testing the comment generator.
            btn_validate (CTkButton): The button for validating the grader report.
            btn_process (CTkButton): The button for processing the report.
            btn_scan_word (CTkButton): The button to start manual scanning of MS Word installation.

        Labels:
            lbl_source (CTkLabel): The label for the source file path.
            lbl_output (CTkLabel): The label for the output folder path.
            lbl_signature (CTkLabel): The label for the signature file path.
            lbl_generate (CTkLabel): The label for the generate options.
            lbl_student_name (CTkLabel): The label for the student name.
            lbl_options (CTkLabel): The label for the options.
            lbl_date (CTkLabel): The label for the report date.
            lbl_progress (CTkLabel): The label for the progress bar.
            lbl_count (CTkLabel): The label for the progress count.
            lbl_status (CTkLabel): The label for the status message.
            lbl_comment_mode (CTkLavel): The label for comment generator mode.

        Textboxes and Entries:
            txt_source_path (CTkEntry): The textbox for the source file path.
            txt_output_path (CTkEntry): The textbox for the output folder path.
            txt_signature_path (CTkEntry): The textbox for the signature file path.
            txt_student_name (CTkEntry): The textbox for the student name.
            txt_status (CTkTextbox): The textbox for the status message.

        Switches:
            switch_autocorrect (CTkSwitch): The switch for enabling/disabling autocorrect.
            switch_date (CTkSwitch): The switch for enabling/disabling the report date.
            switch_force (CTkSwitch): The switch for enabling/disabling force generation.
            switch_pdf (CTkSwitch): The switch for enabling/disabling PDF creation.

        Radio Buttons:
            rdo_generate_all (CTkRadioButton): The radio button for generating all reports.
            rdo_generate_student (CTkRadioButton): The radio button for generating a single report.
            rdo_map_mode (CTkRadioButton): The radio button for comment map mode.
            rdo_ai_mode (CTkRadioButton): The radio button for AI-generated commentmode.

        Date Entry:
            date_report (DateEntry): The date entry for the report date.

        Progress Bar:
            progress_bar (CTkProgressBar): The progress bar for the report generation.

        Variables:
            mode_var (tk.StringVar): The variable for the generate options.
            cgen_mode_var (tk.StringVar): The variable for the comment generator mode.
            autocorrect_var (tk.IntVar): The variable for the autocorrect switch.
            inject_date (tk.IntVar): The variable for the insert date switch.
            force_var (tk.IntVar): The variable for the force generation switch.
            create_pdf (tk.IntVar): The variable for the create PDF switch.

    Methods:
        __init__(self, master, **kwargs): Initializes the ProcessorFrame.
        __browse_file(self): Opens a file dialog for browsing the source file.
        __browse_signature(self): Opens a file dialog for browsing the signature file.
        __save_file(self): Opens a file dialog for saving the output file.
        __opt_all_selected(self): Disables the student name entry when the generate all option is selected.
        __opt_student_selected(self): Enables the student name entry when the generate for student option is selected.
        __toggle_date_entry(self): Enables or disables the date entry field when the insert date option is selected or not.
        __autocorrect_tooltip_message(self): Returns a dynamic tooltip message for the autocorrect switch button.
        __pdf_tooltip_message(self): Returns a dynamic tooltip message for the PDF switch button.
        __process(self): Processes the report based on the selected options.
        __test_source(self): Tests the source file using the comment generator test suite
        __validate(self): Validates the grader report for errors.
        __test_paths(self, source_path, out_path): Tests if the source and output paths are valid.
        __scan_word(self): Scans the device for MS Word installation.
        __on_progress_update(self, current, total, status_message): Updates the progress bar.
        __update_status(self, status_message, clear = False): Updates the status message.
    """

    def __init__(self, master, root, office_version, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

        self.__office_version = office_version
        
        self._grader_report = None

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

        # Comment generator mode
        self.cgen_mode_var = tk.StringVar()
        self.lbl_comment_mode = ctk.CTkLabel(self, text = "Comment Generator Mode:")
        self.rdo_map_mode = ctk.CTkRadioButton(self, text = "Comment Map", variable = self.cgen_mode_var, value = "map", command = self.__map_cgen_mode_selected)
        self.rdo_ai_mode = ctk.CTkRadioButton(self, text = "AI-generated (Experimental Feature)", variable = self.cgen_mode_var, value = "ai", command = self.__ai_cgen_mode_selected)
        self.rdo_map_mode.select()

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
        if not office_version:
            self.switch_pdf.configure(state = tk.DISABLED)
            self.switch_pdf.deselect()

        # Report date
        self.lbl_date = ctk.CTkLabel(self, text = "Report Date:")
        today = date.today()
        self.date_report = tkcal.DateEntry(self, width = 25, 
                                           background = "black", foreground = "white", 
                                           year = today.year, month = today.month, day = today.day, 
                                           font = ("Arial", 12), 
                                           date_pattern = "dd/mm/yyyy",
                                           state = tk.DISABLED)
        
        # Generate button
        self.btn_process = ctk.CTkButton(self, text = "Generate", width = 100, command = self.__process)
        self.btn_test_source = ctk.CTkButton(self, text = "Test Comment Gen", width = 100, fg_color = "grey", command = self.__test_source)
        self.btn_validate = ctk.CTkButton(self, text = "Validate Grader Report", width = 100, fg_color = "grey", command = self.__validate)
        self.btn_scan_word = ctk.CTkButton(self, text = "Re-Scan MS Word", width = 100, fg_color = "grey", command = self.__scan_word)

        # Progress tracker
        self.lbl_progress = ctk.CTkLabel(self, text = "Progress:")
        self.progress_bar = ctk.CTkProgressBar(self, width = 250, mode = "determinate")
        self.lbl_count = ctk.CTkLabel(self, text = "0/0")
        self.progress_bar.set(0)

        # Status section
        self.lbl_status = ctk.CTkLabel(self, text = "Status:")
        self.txt_status = ctk.CTkTextbox(self, width = 300, height = 100, state = tk.DISABLED, wrap = "word")
        self.txt_status.tag_config("warning", foreground = "red")

        # Tooltips
        tooltip_font = ("Arial", 10)
        tooltips = {
            self.btn_browse_source: "Browse for the grader report file.",
            self.btn_browse_output: "Browse for the output folder.",
            self.btn_browse_signature: "Browse for your digitized signature image file.",
            self.btn_scan_word: "Manually scan for MS Word installation. This is already done automatically on startup but you can use this if you have just installed Office, or the startup scan failed to discover your installation.",
            self.rdo_generate_all: "Generate reports for all students.",
            self.rdo_generate_student: "Generate report for a single student. Fill in the student name field to specify the student.",
            self.switch_force: "Enable this to force generate the reports and disregard grader report errors.",
            self.switch_date: "Enable this to insert the date in the report. Please fill in the date field to specify the date to insert.",
            self.txt_student_name: "Enter the name of the student to generate the report for. This is only enabled when the generate for student option is selected.",
            self.date_report: "Select the date to insert in the report. This is only enabled when the insert date option is selected.",
            self.btn_process: "Start generating the reports.",
            self.btn_test_source: "Create a validation list of all possible comment combinations and dumps it into an Excel file",
            self.btn_validate: "Validate the grader report for errors.",
            self.rdo_map_mode: "Use the comment map to generate student comments.",
            self.rdo_ai_mode: "Use AI to generate student comments. Note: This is an experimental feature and may not work as expected. This requires an API key to be set in the configuration file. An API key has been pre-supplied for you but in case the key is not working, please contact the developer or supply your own key. Check the JARS GitHub page for guide on obtaining an API key.",
            self.switch_autocorrect: "Automatically corrects spelling and grammar errors in the generated comments. This requires internet connection as it uses Google's Gemini AI to autocorrect the comments.",
        }

        self.ttip_pdf = tktip.ToolTip(self.switch_pdf, msg = self.__pdf_tooltip_message, font = tooltip_font)

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

        self.lbl_comment_mode.grid(row = 6, column = 0, sticky = tk.W, pady = 2)
        self.rdo_map_mode.grid(row = 6, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.rdo_ai_mode.grid(row = 7, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_options.grid(row = 8, column = 0, sticky = tk.W, pady = 2)
        self.switch_autocorrect.grid(row = 8, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.switch_date.grid(row = 9, column = 1, sticky = tk.EW, padx = 5, pady = 2)

        self.switch_force.grid(row = 10, column = 1, sticky = tk.W, padx = 5, pady = 2)

        self.switch_pdf.grid(row = 11, column = 1, sticky = tk.W, padx = 5, pady = 2)

        self.lbl_date.grid(row = 12, column = 0, sticky = tk.W, pady = 2)
        self.date_report.grid(row = 12, column = 1, sticky = tk.W, padx = 5, pady = 2)

        self.lbl_progress.grid(row = 13, column = 0, sticky = tk.W, pady = 0)
        self.progress_bar.grid(row = 13, column = 1, sticky = tk.W, padx =  5, pady = 0)
        self.lbl_count.grid(row = 13, column = 2, sticky = tk.EW, padx = 2, pady = 0)

        self.lbl_status.grid(row = 14, column = 0, sticky = tk.NW, pady = 0)
        self.txt_status.grid(row = 14, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 0)

        self.btn_test_source.grid(row = 15, column = 0, sticky = tk.EW, padx = 2, pady = (20, 2))
        self.btn_validate.grid(row = 15, column = 1, sticky = tk.W, padx = 2, pady = (20, 2))
        self.btn_process.grid(row = 15, column = 2, sticky = tk.EW, padx = 2, pady = (20, 2))

        self.btn_scan_word.grid(row = 16, column = 0, sticky = tk.EW, padx = 2, pady = (2, 5))

    # UI functions
    def __browse_file(self):
        """Opens a file dialog for browsing the source file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select Grader Report File", defaultextension = ".xlsm", filetypes =[("Microsoft Excel Macro-Enabled Document", "*.xlsm"), ("Microsoft Excel Document", "*.xlsx")])
        self.txt_source_path.delete(0, tk.END)
        self.txt_source_path.insert(0, file_path)

        if os.path.isfile(file_path):
            self._grader_report = grader_report.GraderReport(file_path)

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

    def __map_cgen_mode_selected(self):
        """Enables autocorrect toggle"""
        self.switch_autocorrect.configure(state = tk.NORMAL)

    def __ai_cgen_mode_selected(self):
        """Disables autocorrect toggle"""
        self.switch_autocorrect.configure(state = tk.DISABLED)

    def __toggle_date_entry(self):
        """Enables or disables the date entry field when the insert date option is selected or not."""
        if self.inject_date.get() == 1:
            self.date_report.configure(state = tk.NORMAL)
        else:
            self.date_report.configure(state = tk.DISABLED)
        
    def __pdf_tooltip_message(self):
        """Returns the tooltip message for the PDF switch button."""
        if not self.__office_version:
            return "PDF creation is disabled because Microsoft Office is not installed."
        else:
            return "Enable this to convert the generated reports to PDF."

    def __process(self):
        """
        Processes the report based on the selected options.
        
        It runs the report processor based on the selected options. It will run the report processor
        for all students if the generate all option is selected. It will run the report processor for
        a single student if the generate for student option is selected.
        """
        paths_valid = self.__test_paths()
        if not paths_valid:
            return
        
        output_file_path = self.txt_output_path.get()
        signature_file = self.txt_signature_path.get() if self.txt_signature_path.get() != "" else None
        date = self.date_report.get_date() if self.inject_date.get() == 1 else None

        proc = processor.Generator(output_file_path, self._grader_report, date, signature_file, self.cgen_mode_var.get())

        mode = self.mode_var.get()
        autocorrect = True if self.autocorrect_var.get() == 1 else False
        force = True if self.force_var.get() == 1 else False
        pdf = True if self.create_pdf.get() == 1 else False

        # Pre-checks
        if autocorrect:
            print("Autocorrect enabled!")
            
        if self.cgen_mode_var.get() == "ai":
            print("AI mode selected!")
            print("Checking API key…")

            if config.get_config("genai_api_key") == "" or config.get_config("genai_api_key") is None:
                tk.messagebox.showerror("API Key Not Found", "The AI-generated comment feature requires an API key. Please enter the API key in the configuration file and restart the program.")
                self.rdo_ai_mode.deselect()
                self.rdo_map_mode.select()
                self.rdo_ai_mode.configure(state = tk.DISABLED)
                self.__update_status("Aborting report generation!")
                return

            tk.messagebox.showwarning("Experimental Feature", "The AI-generated comment feature is an experimental feature and may not work as expected. Please use with caution.")
        
        # Start report generation
        self.__update_status("Starting report generation…", clear = True)
        
        if mode == "all":
            proc.generate_all(callback = self.__on_progress_update, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
        elif mode == "student":
            student_name = self.txt_student_name.get()
            self.__on_progress_update(0, 1, f"Generating report for {student_name}…")
            proc.generate_for_student(student_name = student_name, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
            output_file_path = f"{output_file_path}/{student_name}.docx"

        # Post-operation
        self.lbl_count.configure(text = "Done!")
        self.__update_status("Report generation completed successfully.")

        OutputDialog(master = self.root, title = "Report Generation Complete", file_path = output_file_path, content = "Report generation completed successfully.")

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
        
        self.__update_status("Running comment generator test suite. Please wait…", clear = True)
        
        cgen_test.run(source_file_path = source_file, output_file_path = output_file_path, CommentGenerator = CommentGenerator)
        
        OutputDialog(master = self.root, title = "Comment Generator Test Complete", file_path = output_file_path, content = "Comment generator test completed successfully.")
        self.__update_status("Comment generator test completed successfully.")

    def __validate(self):
        """
        Validates the grader report for errors.
        It runs the grader report validator and displays the result in a message box.
        
        This function does not return anything. It displays a message box with the result of the validation.
        """
        source_file = self.txt_source_path.get()
        
        if not os.path.isfile(source_file):
            tk.messagebox.showerror("Error", "Please select a valid source file.")
            return
        
        gr = grader_report.GraderReport(source_file, skip_validation = True)
        valid = gr.validate(callback = self.__update_status)
        if not valid:
            tk.messagebox.showerror("Invalid grader report", "Grader report is invalid. Check console/terminal for details.")
        else:
            tk.messagebox.showinfo("Valid grader report", "Grader report is valid. You're good to go!")
            
        self.lbl_count.configure(text = "Done!")

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
    
    def __scan_word(self):
        """
        Scans the device for MS Word installation.
        
        This function does not return anything. It updates the GUI and enables the PDF switch if MS Word is found.
        """
        import wmi
        from win32com.client import Dispatch

        if self.__office_version:
            tk.messagebox.showinfo("Microsoft Office Detected", "Microsoft Office has already been detected. No need to scan.")
            return

        print("> Manually Scanning for MS Word installation…")

        # Check if MS Word is running
        print("  Checking if an MS Word instance is running…")
        word_is_open = False
        c = wmi.WMI()
        for process in c.Win32_Process():
            if process.Name == "WINWORD.EXE":
                word_is_open = True
        
        word_is_open = False
        
        # Check if Microsoft Office is installed
        try:
            print("  Checking if Microsoft Office is installed…")
            word = Dispatch("Word.Application")
            office_version = word.Version
            if not word_is_open:
                word.Quit()
            print(f"  Microsoft Office {office_version} detected.")
        except:
            office_version = None
            print("  Microsoft Office not detected.")

        # Set GUI
        if office_version:
            self.__office_version = office_version
            self.switch_pdf.configure(state = tk.NORMAL)
            self.__update_status("Microsoft Office detected. PDF creation is enabled.")
            print("  Microsoft Office detected. PDF creation is enabled.")

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

        self.__update_status(status_message)

    def __update_status(self, status_message, clear = False):
        """
        Updates the status message.
        Can be used as a callback function for the report processor.

        Args:
            status_message (str): The status message to display.
        """
        tag = None

        self.txt_status.configure(state = tk.NORMAL)
        
        if clear:
            self.txt_status.delete("1.0", tk.END)

        if "warning" in status_message.lower():
            tag = "warning"

        self.txt_status.insert(tk.END, f"{status_message}\n", tags = tag)
        self.txt_status.configure(state = tk.DISABLED)
        self.txt_status.see(tk.END)

        tk.Misc.update_idletasks(self)