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

        Date Entry:
            date_report (DateEntry): The date entry for the report date.

        Progress Bar:
            progress_bar (CTkProgressBar): The progress bar for the report generation.

        Variables:
            mode_var (tk.StringVar): The variable for the generate options.
            autocorrect_var (tk.IntVar): The variable for the autocorrect switch.
            inject_date (tk.IntVar): The variable for the insert date switch.
            force_var (tk.IntVar): The variable for the force generation switch.
            create_pdf (tk.IntVar): The variable for the create PDF switch.

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

    def __init__(self, master, root, office_version, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

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
        self.rdo_map_mode = ctk.CTkRadioButton(self, text = "Comment Map", variable = self.cgen_mode_var, value = "map")
        self.rdo_ai_mode = ctk.CTkRadioButton(self, text = "AI-generated (Experimental Feature)", variable = self.cgen_mode_var, value = "ai")
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
            tktip.ToolTip(self.switch_pdf, "PDF creation is disabled because Microsoft Office is not installed.", font = ("Arial", 10))
        else:
            tktip.ToolTip(self.switch_pdf, "Enable this to convert the generated reports to PDF.", font = ("Arial", 10))

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
            self.rdo_generate_all: "Generate reports for all students.",
            self.rdo_generate_student: "Generate report for a single student. Fill in the student name field to specify the student.",
            self.switch_force: "Enable this to force generate the reports and disregard grader report errors.",
            self.switch_date: "Enable this to insert the date in the report. Please fill in the date field to specify the date to insert.",
            self.txt_student_name: "Enter the name of the student to generate the report for. This is only enabled when the generate for student option is selected.",
            self.date_report: "Select the date to insert in the report. This is only enabled when the insert date option is selected.",
            self.btn_process: "Start generating the reports.",
            self.btn_test_source: "Create a validation list of all possible comment combinations and dumps it into an Excel file",
            self.btn_validate: "Validate the grader report for errors."
        }

        self.__autocorrect_disabled = False
        self.ttip_switch = tktip.ToolTip(self.switch_autocorrect, msg = self.__autocorrect_tooltip_message, font = tooltip_font)

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

    def __autocorrect_tooltip_message(self):
        """Returns the tooltip message for the autocorrect switch button."""
        if self.__autocorrect_disabled:
            return "Autocorrect has been disabled for this session because Java is not installed."
        else:
            return "Enable this to autocorrect the comments using Language Tool. Note: This requires Java to be installed."

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

        if autocorrect:
            java_exists = ltm.check_java()
            ltm_exists = ltm.check_package()
            install_java = False
            download_lt = False

            if not java_exists:
                install_java = tk.messagebox.askyesno("Java Not Found", 
                                "To use the autocorrect feature, you need to install Java. Do you want to download and install it now?")
                if install_java:
                    tk.messagebox.showinfo("Java Download", "JARS will be closed and you will be redirected to the Java download page on your browser. Please install Java, restart the software, and try again.")
                    java_url = "https://www.java.com/en/download/"
                    import webbrowser, sys
                    webbrowser.open_new(java_url)
                    sys.exit()

            if not ltm_exists and (install_java or java_exists):
                download_lt = tk.messagebox.askokcancel("Language Tool Package Not Found", 
                                "To use the autocorrect feature, you need to install the Language Tool package. Do you want to download and install it now (~200 MB)? Note: This Requires Java to be installed")
                
            if (not download_lt and not ltm_exists) or (not install_java and not java_exists):
                tk.messagebox.showinfo("Autocorrect Disabled", "For this session, autocorrect has been disabled because Java is not installed and/or Language Tool package is not found. You can still use the software without the autocorrect feature.")
                self.switch_autocorrect.deselect()
                self.switch_autocorrect.configure(state = tk.DISABLED)
                self.__autocorrect_disabled = True
                return
            
        self.__update_status("Starting report generation…", clear = True)
        
        if mode == "all":
            proc.generate_all(callback = self.__on_progress_update, mode = self.cgen_mode_var.get(), autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
        elif mode == "student":
            student_name = self.txt_student_name.get()
            self.__on_progress_update(0, 1, f"Generating report for {student_name}…")
            proc.generate_for_student(student_name = student_name, mode = self.cgen_mode_var.get(), autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
            output_file_path = f"{output_file_path}/{student_name}.docx"

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
        
        cgen_test.run(source_file_path = source_file, output_file_path = output_file_path, ltm = ltm, CommentGenerator = CommentGenerator)
        
        OutputDialog(master = self.root, title = "Comment Generator Test Complete", file_path = output_file_path, content = "Comment generator test completed successfully.")
        self.__update_status("Comment generator test completed successfully.")

    def __validate(self):
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