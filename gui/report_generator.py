import os
import threading
from datetime import date, datetime, timedelta

import customtkinter as ctk
import tkinter as tk
import tktooltip as tktip
import tkcalendar as tkcal
from tkinter import ttk
from windows_toasts import InteractableWindowsToaster, Toast, ToastAudio, AudioSource, ToastButton

import config
import components.report_generator.semester_report as processor
import components.common.grader_report as grader_report
import components.report_generator.comment_generator_test as cgen_test
import components.utility as util
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
        self.processor_frame.grid(padx = 10, pady = 10)
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
        MENUBAR SETUP
        """
        self.menubar = tk.Menu(self)

        self.menu_utility = tk.Menu(self.menubar, tearoff = False)
        self.menu_utility.add_command(label = "Test Grader Report Comment Mapper", command = self.__test_source)
        self.menu_utility.add_command(label = "Validate Grader Report", command = self.__validate)
        self.menu_utility.add_separator()
        self.menu_utility.add_command(label = "Scan for MS Word Installation", command = self.__scan_word)

        self.menu_preference = tk.Menu(self.menubar, tearoff = False)
        self.save_signature_path_var = tk.BooleanVar()
        self.submenu_save_signature_path = self.menu_preference.add_checkbutton(label = "Save signature path", variable = self.save_signature_path_var, command = self.__save_signature_path, onvalue = 1, offvalue = 0)
        self.always_on_pdf_var = tk.BooleanVar()
        self.submenu_always_on_pdf = self.menu_preference.add_checkbutton(label = "Always create PDF", variable = self.always_on_pdf_var, command = self.__always_on_pdf, onvalue = 1, offvalue = 0)
        self.menu_preference.add_separator()
        self.menu_preference.add_command(label = "Settings…", command = self.__open_configurator, accelerator = "Ctrl+P")

        self.menubar.add_cascade(label = "Utilities", menu = self.menu_utility)
        self.menubar.add_cascade(label = "Preferences", menu = self.menu_preference)
        self.master.config(menu = self.menubar)

        # Menubar Shortcut Bindings
        self.master.bind_all("<Control-p>", lambda event: self.__open_configurator())
        self.master.bind_all("<Control-P>", lambda event: self.__open_configurator())

        # Set dynamic menu item states
        if config.get_config("signature_path") == "" or config.get_config("signature_path") is None:
            self.save_signature_path_var.set(False)
        else:
            self.save_signature_path_var.set(True)

        self.always_on_pdf_var.set(config.get_config("always_create_pdf"))

        """
        WIDGETS SETUP
        """
        # Grader report file path
        self.lbl_source = ctk.CTkLabel(self, text = "Grader Report File:")
        self.txt_source_path = ctk.CTkEntry(self, width = 450)
        self.btn_browse_source = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__browse_file)

        # Output file path
        self.lbl_output = ctk.CTkLabel(self, text = "Output Folder:")
        self.txt_output_path = ctk.CTkEntry(self, width = 450)
        self.btn_browse_output = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__save_file)

        # Signature file path
        self.lbl_signature = ctk.CTkLabel(self, text = "Signature File:")
        self.txt_signature_path = ctk.CTkEntry(self, width = 450, placeholder_text = "Leave blank to not include signature")
        self.btn_browse_signature = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__browse_signature)

        # Generate all or generate for student
        self.mode_var = tk.StringVar()
        self.lbl_generate = ctk.CTkLabel(self, text = "Generate:")
        self.rdo_generate_all = ctk.CTkRadioButton(self, text = "All", variable = self.mode_var, value = "all", command = self.__opt_all_selected)
        self.rdo_generate_student = ctk.CTkRadioButton(self, text = "Student", variable = self.mode_var, value = "student", command = self.__opt_student_selected)

        # Student name
        self.lbl_student_name = ctk.CTkLabel(self, text = "Student Name:")
        self.txt_student_name = ctk.CTkEntry(self, width = 450)

        # Comment generator mode
        self.cgen_mode_var = tk.StringVar()
        self.lbl_comment_mode = ctk.CTkLabel(self, text = "Comment Generator Mode:")
        self.rdo_map_mode = ctk.CTkRadioButton(self, text = "Comment Map", variable = self.cgen_mode_var, value = "map", command = self.__map_cgen_mode_selected)
        self.rdo_ai_mode = ctk.CTkRadioButton(self, text = "AI-generated", variable = self.cgen_mode_var, value = "ai", command = self.__ai_cgen_mode_selected)

        # Test AI API key button
        self.btn_test_api_key = ctk.CTkButton(self, text = "Test API Key", width = 100, fg_color = "grey", command = self.__test_api_key)

        # Options section
        self.lbl_options = ctk.CTkLabel(self, text = "Options:")

        # Autocorrect switch button
        self.autocorrect_var = tk.IntVar()
        self.switch_autocorrect = ctk.CTkSwitch(self, text = "Autocorrect", variable = self.autocorrect_var, onvalue = 1, offvalue = 0)

        # Force generation switch button
        self.force_var = tk.IntVar()
        self.switch_force = ctk.CTkSwitch(self, text = "Force Generate", variable = self.force_var, onvalue = 1, offvalue = 0)

        # Insert date switch button
        self.inject_date = tk.IntVar()
        self.switch_date = ctk.CTkSwitch(self, text = "Insert Date. Report Date      ⟶   ", command = self.__toggle_date_entry, variable = self.inject_date, onvalue = 1, offvalue = 0)
        
        # Report date
        today = date.today()
        self.date_report = tkcal.DateEntry(self, width = 25, 
                                           background = "black", foreground = "white", 
                                           year = today.year, month = today.month, day = today.day, 
                                           font = ("Arial", 12), 
                                           date_pattern = "dd/mm/yyyy",
                                           state = tk.DISABLED)

        # Create PDF switch button
        self.create_pdf = tk.IntVar()
        self.switch_pdf = ctk.CTkSwitch(self, text = "Create PDF", variable = self.create_pdf, onvalue = 1, offvalue = 0)

        # Progress tracker
        self.lbl_progress = ctk.CTkLabel(self, text = "Progress:")
        self.progress_bar = ctk.CTkProgressBar(self, width = 450, mode = "determinate")
        self.lbl_count = ctk.CTkLabel(self, text = "0/0")
        self.progress_bar.set(0)

        # Status section
        self.lbl_status = ctk.CTkLabel(self, text = "Status:")
        self.txt_status = ctk.CTkTextbox(self, width = 600, height = 150, state = tk.DISABLED, wrap = "word")
        self.txt_status.tag_config("warning", foreground = "red")

        # Generate button
        self.btn_process = ctk.CTkButton(self, text = "Generate", width = 100, command = self.__threaded_process)
        self.btn_validate = ctk.CTkButton(self, text = "Validate Grader Report", width = 150, fg_color = "grey", command = self.__validate)
        self.btn_configure = ctk.CTkButton(self, text = "Settings…", width = 150, fg_color = "purple", command = self.__open_configurator)

        # Tree View
        self.lbl_treeview = ctk.CTkLabel(self, text = "Grader Report Explorer")
        self.lbl_tv_search = ctk.CTkLabel(self, text = "Search:", width = 50)
        self.txt_tv_search = ctk.CTkEntry(self, width = 300)
        style = ttk.Style()
        
        if ctk.get_appearance_mode() == "Dark":
            style.theme_use("alt")
            style.configure("JARS.Treeview", rowheight = 35, background = "gray24", fieldbackground = "gray14", foreground = "white")
        else:
            style.configure("JARS.Treeview", rowheight = 35)

        self.treeview = ttk.Treeview(self, columns = ("value"), style = "JARS.Treeview")
        self.treeview.column("#0", width = 300)
        self.treeview.column("value", width = 300)
        self.treeview.heading("#0", text = "Key")
        self.treeview.heading("value", text = "Value")

        self.vsb_treeview = ttk.Scrollbar(self, orient = "vertical", command = self.treeview.yview)

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
            self.btn_validate: "Validate the grader report for errors.",
            self.rdo_map_mode: "Use the comment map to generate student comments.",
            self.rdo_ai_mode: "Use AI to generate student comments. Note: This is an experimental feature and may not work as expected. This requires an API key to be set in the configuration file. An API key has been pre-supplied for you but in case the key is not working, please contact the developer or supply your own key. Check the JARS GitHub page for guide on obtaining an API key.",
            self.btn_test_api_key: "Test the Google Gemini AI API key set in the configuration.",
            self.btn_configure: "Open the configurator to set API keys and other settings.",
            self.switch_autocorrect: "Automatically corrects spelling and grammar errors in the generated comments. This requires internet connection as it uses Google's Gemini AI to autocorrect the comments.",
        }

        self.ttip_pdf = tktip.ToolTip(self.switch_pdf, msg = self.__pdf_tooltip_message, font = tooltip_font)

        for widget, message in tooltips.items():
            tktip.ToolTip(widget, message, font = tooltip_font)
        
        """
        GUI LAYOUTING
        """
        self.columnconfigure(1, weight = 1)
        self.rowconfigure(1, weight = 1)

        self.lbl_source.grid(row = 0, column = 0, sticky = tk.W, pady = 2)
        self.txt_source_path.grid(row = 0, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 2)
        self.btn_browse_source.grid(row = 0, column = 3, sticky = tk.EW, padx = 2, pady = 2)
        self.lbl_treeview.grid(row = 0, column = 4, columnspan = 3, sticky = tk.EW, padx = 10, pady = 2)

        self.lbl_output.grid(row = 1, column = 0, sticky = tk.W, pady = 2)
        self.txt_output_path.grid(row = 1, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 2)
        self.btn_browse_output.grid(row = 1, column = 3, sticky = tk.EW, padx = 2, pady = 2)
        self.lbl_tv_search.grid(row = 1, column = 4, sticky = tk.W, padx = (10, 2), pady = 2)
        self.txt_tv_search.grid(row = 1, column = 5, columnspan = 2, sticky = tk.EW, padx = (2, 10), pady = 2)

        self.lbl_signature.grid(row = 2, column = 0, sticky = tk.W, pady = 2)
        self.txt_signature_path.grid(row = 2, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 2)
        self.btn_browse_signature.grid(row = 2, column = 3, sticky = tk.EW, padx = 2, pady = 2)
        self.treeview.grid(row = 2, column = 4, rowspan = 10, columnspan = 2, sticky = tk.NSEW, padx = (10, 0), pady = 2)
        self.vsb_treeview.grid(row = 2, column = 6, rowspan = 10, sticky = tk.NS, padx = (0, 10), pady = 2)

        self.lbl_generate.grid(row = 3, column = 0, sticky = tk.W, pady = 2)
        self.rdo_generate_all.grid(row = 3, column = 1, sticky = tk.W, padx =  5, pady = 2)
        
        self.rdo_generate_student.grid(row = 3, column = 2, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_student_name.grid(row = 4, column = 0, sticky = tk.W, pady = 2)
        self.txt_student_name.grid(row = 4, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 2)

        self.lbl_comment_mode.grid(row = 5, column = 0, sticky = tk.W, pady = 2)
        self.rdo_map_mode.grid(row = 5, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.rdo_ai_mode.grid(row = 5, column = 2, sticky = tk.W, padx =  5, pady = 2)
        self.btn_test_api_key.grid(row = 5, column = 3, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_options.grid(row = 6, column = 0, sticky = tk.W, pady = 2)
        self.switch_autocorrect.grid(row = 6, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.switch_force.grid(row = 6, column = 2, sticky = tk.W, padx = 5, pady = 2)

        self.switch_date.grid(row = 7, column = 1, sticky = tk.EW, padx = 5, pady = 2)
        self.date_report.grid(row = 7, column = 2, sticky = tk.EW, padx = (5, 10), pady = 2)

        self.switch_pdf.grid(row = 8, column = 1, sticky = tk.W, padx = 5, pady = 2)
        
        self.lbl_progress.grid(row = 9, column = 0, sticky = tk.W, pady = 0)
        self.progress_bar.grid(row = 9, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 0)
        self.lbl_count.grid(row = 9, column = 3, sticky = tk.EW, padx = 2, pady = 0)

        self.lbl_status.grid(row = 10, column = 0, sticky = tk.NW, pady = 0)
        self.txt_status.grid(row = 10, column = 1, columnspan = 3, sticky = tk.EW, padx =  5, pady = 0)

        self.btn_configure.grid(row = 11, column = 0, sticky = tk.EW, padx = 2, pady = (20, 2))
        self.btn_validate.grid(row = 11, column = 1, sticky = tk.EW, padx = 2, pady = (20, 2))
        self.btn_process.grid(row = 11, column = 3, sticky = tk.EW, padx = 2, pady = (20, 2))

        """
        POST LAYOUTING SETUP
        """
        # Populate signature path
        if config.get_config("signature_path") != "" and config.get_config("signature_path") is not None:
            self.txt_signature_path.insert(0, config.get_config("signature_path"))

        # Select generate all by default and disable student name entry
        self.rdo_generate_all.select()
        self.txt_student_name.insert(0, "This field is disabled because you selected to generate all reports.")
        self.txt_student_name.configure(state = tk.DISABLED, text_color = "grey")

        # Select AI mode by default and disable autocorrect switch
        self.rdo_ai_mode.select()
        self.switch_autocorrect.configure(state = tk.DISABLED)

        # Always on PDF switch
        if self.always_on_pdf_var.get():
            self.switch_pdf.select()

        # Check if MS Word is installed and disable PDF switch if not found
        if not office_version:
            self.switch_pdf.configure(state = tk.DISABLED)
            self.switch_pdf.deselect()

        # Bind search function to search entry
        self.txt_tv_search.bind("<Return>", lambda event: self.__search_treeview())
        self.txt_tv_search.bind("<KP_Enter>", lambda event: self.__search_treeview())

        # Bind scrollbars to treeview
        self.treeview.configure(yscrollcommand = self.vsb_treeview.set)

    # UI functions
    def __browse_file(self):
        """Opens a file dialog for browsing the source file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select Grader Report File", defaultextension = ".xlsm", filetypes =[("Microsoft Excel Macro-Enabled Document", "*.xlsm"), ("Microsoft Excel Document", "*.xlsx")])
        self.txt_source_path.delete(0, tk.END)
        self.txt_source_path.insert(0, file_path)

        if os.path.isfile(file_path):
            self._grader_report = grader_report.GraderReport(file_path, skip_validation = True, callback = self.__update_status)
            if not self._grader_report:
                tk.messagebox.showerror("Error", "An error occurred while loading the grader report. Check console/terminal for details.")
                return

            self.__update_status("Validating grader report…", clear = True)
            valid = self._grader_report.validate(callback = self.__update_status)
            if not valid:
                self.__update_status("\nWARNING: Grader report is invalid. Check information above for details.")
                self.treeview.delete(*self.treeview.get_children())
                self.treeview.insert("", tk.END, text = "Grader Report Invalid.")
            else:
                self.__update_status("Grader report is valid. You're good to go!")
                self.__populate_treeview()

    def __populate_treeview(self):
        """Populates the tree view with the grader report data."""
        self.treeview.delete(*self.treeview.get_children()) # Clear treeview
        
        self.treeview.insert("", tk.END, text = "School Year", values = (self._grader_report.get_course_info("School Year")))
        self.treeview.insert("", tk.END, text = "Semester", values = (self._grader_report.get_course_info("Semester")))
        self.treeview.insert("", tk.END, text = "Subject", values = (self._grader_report.get_course_info("Subject")))
        self.treeview.insert("", tk.END, text = "Grade", values = (self._grader_report.get_course_info("Grade")))

        self.treeview.insert("", tk.END, text = "Total Students", values = (self._grader_report.count_students()))

        # Show student list from dataframe
        tv_student_list = self.treeview.insert("", tk.END, text = "Student List:")
        for index, row in self._grader_report.students.iterrows():
            # Generic Display
            student = self.treeview.insert(tv_student_list, tk.END, text = index)
            self.treeview.insert(student, tk.END, text = "Short Name", values = (row["Short Name"]))
            self.treeview.insert(student, tk.END, text = "Gender", values = (row["Gender"]))
            self.treeview.insert(student, tk.END, text = "Grade (Numeric)", values = (self._grader_report.get_final_grade(index, "Final Score")))
            self.treeview.insert(student, tk.END, text = "Grade (Letter)", values = (self._grader_report.get_final_grade(index, "Letter Grade")))
            
            # PD Display
            pd = self.treeview.insert(student, tk.END, text = "Personal Development")
            for i, item in enumerate(self._grader_report.data_pd.columns):
                self.treeview.insert(pd, tk.END, text = item, values = (self._grader_report.get_grade_pd(index, item)))

            # SNA Display
            sna = self.treeview.insert(student, tk.END, text = "Skills and Assessment")
            for i, item in enumerate(self._grader_report.data_sna.columns):
                self.treeview.insert(sna, tk.END, text = item, values = (self._grader_report.get_grade_sna(index, item)))

    def __search_treeview(self):
        """Searches the tree view for the specified text."""
        search_text = self.txt_tv_search.get().lower()
        if search_text == "":
            return

        selections = []
        for child in self.treeview.get_children():
            if search_text in self.treeview.item(child)["text"].lower():
                selections.append(child)
            else:
                for subchild in self.treeview.get_children(child):
                    if search_text in self.treeview.item(subchild)["text"].lower():
                        selections.append(subchild)

        if selections:
            self.treeview.selection_set(selections)

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
        self.txt_student_name.delete(0, tk.END)
        self.txt_student_name.insert(0, "This field is disabled because you selected to generate all reports.")
        self.txt_student_name.configure(state = tk.DISABLED, text_color = "grey")

    def __opt_student_selected(self):
        """Enables the student name entry when the generate for student option is selected."""
        self.txt_student_name.configure(state = tk.NORMAL, text_color = ("black", "white"))
        self.txt_student_name.delete(0, tk.END)
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
        
    def __check_threaded_process(self, process_thread):
        """
        Watches the progress of the threaded report generation process.
        When the process is completed, it will re-enable the process button and show a completion message.
        """
        if process_thread.is_alive():
            self.master.after(1000, self.__check_threaded_process, process_thread)
        else:
            self.btn_process.configure(state = tk.NORMAL)
            self.btn_validate.configure(state = tk.NORMAL)

            # Post-operation
            self.lbl_count.configure(text = "Done!")
            self.__update_status("Report generation completed successfully.")

            # Show completion dialog
            output_file_path = self.txt_output_path.get()
            OutputDialog(master = self.root, 
                         title = "Report Generation Complete", 
                         file_path = output_file_path, 
                         content = "Report generation completed successfully.")
            
            # Toast Notification
            toaster = InteractableWindowsToaster("JARS Report Generator", "jac.acreportingsystem.crep")
            
            toast_finish = Toast()
            toast_finish.text_fields = ["Report generation completed successfully."]
            toast_finish.audio = ToastAudio(AudioSource.Call7)
            toast_finish.AddAction(ToastButton("Open Folder", "open_folder"))
            toast_finish.AddAction(ToastButton("Dismiss", "dismiss"))
            toast_finish.on_activated = self.__toast_button_click
            toast_finish.expiration_time = datetime.now() + timedelta(days = 1)
            
            toaster.show_toast(toast_finish)

    def __threaded_process(self):
        """Starts the report generation in a separate thread."""
        process_thread = threading.Thread(target = self.__process)
        self.btn_process.configure(state = tk.DISABLED)
        self.btn_validate.configure(state = tk.DISABLED)
        process_thread.start()
        self.master.after(1000, self.__check_threaded_process, process_thread)

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
            else:
                print("API key found!\nValidating API key with Google…")
                genai_api_key_validated = self.__test_api_key(suppress_dialog = True)

            if not genai_api_key_validated:
                tk.messagebox.showerror("Invalid API Key", "The API key provided is invalid. Please check the key in the configuration file and try again.")
                self.__update_status("Warning: Invalid API key. Please check the configuration file or use the configurator to set a valid API key.")
                self.__update_status("Aborting report generation!")
                return

            tk.messagebox.showwarning("Attention: Internet Required", "The AI-based comment generator utilizes Google's Gemini AI. Hence, you are required to connect to the internet when using this feature.")
        
        # Start report generation
        self.__update_status("Starting report generation…", clear = True)
        
        if mode == "all":
            proc.generate_all(callback = self.__on_progress_update, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
        elif mode == "student":
            student_name = self.txt_student_name.get()
            self.__on_progress_update(0, 1, f"Generating report for {student_name}…")
            proc.generate_for_student(student_name = student_name, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
            output_file_path = f"{output_file_path}/{student_name}.docx"

    def __toast_button_click(self, toastEvent):
        """Event handler for the toast button click."""
        if toastEvent.arguments == "open_folder":
            os.startfile(self.txt_output_path.get())
        elif toastEvent.arguments == "dismiss":
            pass

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
        
        gr = grader_report.GraderReport(source_file, skip_validation = True, callback = self.__update_status)
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

    def __test_api_key(self, suppress_dialog = False):
        """
        Tests the GenAI API key.
        
        This function does not return anything. It displays a message box with the result of the API key validation.
        """
        valid = False
        output_text = "API key is invalid. Please check the configuration file or use the configurator to set a valid API key. Make sure you are connected to the internet."
        
        print("Testing API key…")
        self.__update_status("Testing API key…", clear = True)

        if util.validate_genai_api_key():
            valid = True
            output_text = "API key is valid. You're good to go!"
            print("> Valid.")
        else:
            print("> Invalid.")

        self.__update_status(output_text)
        if not suppress_dialog:
            tk.messagebox.showinfo("API Key Test", output_text)

        return valid

    def __open_configurator(self):
        """Opens the configuration window."""
        from gui.configurator import ConfiguratorWindow
        ConfiguratorWindow(master = self.root)

    def __save_signature_path(self):
        """Saves the signature path to the configuration file."""
        if self.save_signature_path_var.get():
            new_config = config.set_config("signature_path", self.txt_signature_path.get())
            config.save_config(new_config)
            self.__update_status("Signature path saved to configuration file.")
        else:
            new_config = config.set_config("signature_path", "")
            config.save_config(new_config)
            self.__update_status("Signature path removed from configuration file.")

    def __always_on_pdf(self):
        """Saves the always create PDF setting to the configuration file."""
        new_config = config.set_config("always_create_pdf", self.always_on_pdf_var.get())
        config.save_config(new_config)

        if self.always_on_pdf_var.get():
            self.switch_pdf.select()
        else:
            self.switch_pdf.deselect()

        self.__update_status("Always create PDF setting saved to configuration file.")

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