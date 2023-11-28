import customtkinter as ctk
import tkinter as tk

import processor.semester_report as processor
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
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")

        # Define widgets
        # Grader report file path
        self.lbl_source = ctk.CTkLabel(self, text = "Grader Report File:")
        self.txt_source_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_source = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__browse_file)

        # Output file path
        self.lbl_output = ctk.CTkLabel(self, text = "Output Folder:")
        self.txt_output_path = ctk.CTkEntry(self, width = 250)
        self.btn_browse_output = ctk.CTkButton(self, text = "Browse…", width = 100, command = self.__save_file)

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

        # Generate button
        self.btn_process = ctk.CTkButton(self, text = "Generate", width = 100, command = self.__process)
        self.btn_test_source = ctk.CTkButton(self, text = "Test Comment Gen", width = 100, command = self.__test_source)
        self.btn_test_source.configure(fg_color = "grey")

        # Progress tracker
        self.lbl_progress = ctk.CTkLabel(self, text = "Progress:")
        self.progress_bar = ctk.CTkProgressBar(self, width = 250, mode = "determinate")
        self.lbl_count = ctk.CTkLabel(self, text = "0/0")
        self.progress_bar.set(0)

        # Status section
        self.lbl_status = ctk.CTkLabel(self, text = "Status:")
        self.lbl_status_text = ctk.CTkLabel(self, width = 300, justify = "left", anchor = tk.W, text = "Idle")

        # Layout
        self.lbl_source.grid(row = 0, column = 0, sticky = tk.W, pady = 2)
        self.txt_source_path.grid(row = 0, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_source.grid(row = 0, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_output.grid(row = 1, column = 0, sticky = tk.W, pady = 2)
        self.txt_output_path.grid(row = 1, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.btn_browse_output.grid(row = 1, column = 2, sticky = tk.EW, padx = 2, pady = 2)

        self.lbl_generate.grid(row = 2, column = 0, sticky = tk.W, pady = 2)
        self.rdo_generate_all.grid(row = 2, column = 1, sticky = tk.W, padx =  5, pady = 2)
        self.rdo_generate_student.grid(row = 3, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_student_name.grid(row = 4, column = 0, sticky = tk.W, pady = 2)
        self.txt_student_name.grid(row = 4, column = 1, sticky = tk.W, padx =  5, pady = 2)
        
        self.lbl_options.grid(row = 5, column = 0, sticky = tk.W, pady = 2)
        self.switch_autocorrect.grid(row = 5, column = 1, sticky = tk.W, padx =  5, pady = 2)

        self.lbl_progress.grid(row = 6, column = 0, sticky = tk.W, pady = 0)
        self.progress_bar.grid(row = 6, column = 1, sticky = tk.W, padx =  5, pady = 0)
        self.lbl_count.grid(row = 6, column = 2, sticky = tk.EW, padx = 2, pady = 0)

        self.lbl_status.grid(row = 7, column = 0, sticky = tk.W, pady = 0)
        self.lbl_status_text.grid(row = 7, column = 1, columnspan = 2, sticky = tk.EW, padx =  5, pady = 0)

        self.btn_test_source.grid(row = 8, column = 0, sticky = tk.EW, padx = 2, pady = (20, 2))
        self.btn_process.grid(row = 8, column = 2, sticky = tk.EW, padx = 2, pady = (20, 2))
        

    # UI functions
    def __browse_file(self):
        """Opens a file dialog for browsing the source file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select Grader Report File", defaultextension = ".xlsm", filetypes =[("Microsoft Excel Macro-Enabled Document", "*.xlsm"), ("Microsoft Excel Document", "*.xlsx")])
        self.txt_source_path.delete(0, tk.END)
        self.txt_source_path.insert(0, file_path)

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

    def __process(self):
        """Processes the report based on the selected options."""
        source_file = self.txt_source_path.get()
        output_file_path = self.txt_output_path.get()
        mode = self.mode_var.get()
        proc = processor.Generator(source_file, output_file_path)

        autocorrect = True if self.autocorrect_var.get() == 1 else False
        
        if mode == "all":
            proc.generate_all(callback = self.__on_progress_update, autocorrect = autocorrect)
        elif mode == "student":
            student_name = self.txt_student_name.get()
            self.__on_progress_update(0, 1, f"Generating report for {student_name}…")
            proc.generate_for_student(student_name = student_name, autocorrect = autocorrect)
            self.__on_progress_update(1, 1, f"Done!")
            output_file_path = f"{output_file_path}/{student_name}.docx"

        self.lbl_count.configure(text = "Done!")
        self.lbl_status_text.configure(text = "Report generation completed successfully.")
        OutputDialog(master = self.master, title = "Report Generation Complete", file_path = output_file_path, content = "Report generation completed successfully.")

    def __test_source(self):
        """Tests the source file using the comment generator test suite"""
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

    def __on_progress_update(self, current, total, status_message):
        """Updates the progress bar."""
        progress = current / total
        
        self.progress_bar.set(progress)
        self.lbl_count.configure(text = f"{current}/{total}")

        self.lbl_status_text.configure(text = status_message)

        tk.Misc.update_idletasks(self)