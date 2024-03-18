import customtkinter as ctk
import tkinter as tk

import components.common.integrity as integrity

class InManageVerifierWindow(ctk.CTkToplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        # Window setup
        self.title("JARS InManage")
        self.resizable(False, False)

        # Frame setup
        self.inmanage_frame = InManageVerifierFrame(master = self, root = master)
        self.inmanage_frame.pack(fill = tk.BOTH, expand = True, padx = 10, pady = 10)
        master.eval(f"tk::PlaceWindow {self} center")

        # Hide master window
        self.master.withdraw()
        self.bind("<Destroy>", self.__on_destroy)

    def __on_destroy(self, event):
        """Shows the master window when this window is destroyed."""
        if event.widget == self:
            self.master.deiconify()

class InManageVerifierFrame(ctk.CTkFrame):
    """
    A custom frame for the InManage Verifier application.

    This frame contains various widgets for selecting a file to verify and verifying it.
    Also displays the result of the verification.

    Attributes:
        Buttons:
            btn_file (CTkButton): The button for browsing the file to verify.
            btn_verify (CTkButton): The button for verifying the selected file.

        Labels:
            lbl_title (CTkLabel): The application window title.
            lbl_file (CTkLabel): The label for the file selection.
            lbl_output (CTkLabel): The label for the output text box.
            lbl_info (CTkLabel): The label for the verification result.

        Textboxes and Entries:
            txt_file (CTkEntry): The text entry for the file to verify.
            txt_output (CTkTextbox): The text box for displaying the verification result.
    """
    def __init__(self, master, root, **kwargs):
        super().__init__(master, **kwargs, fg_color = "transparent")
        self.root = root

        """
        WIDGETS SETUP
        """
        # Title
        self.lbl_title = ctk.CTkLabel(self, text = "InManage | Report Integrity Verifier", font = ("Arial", 20, "bold"))

        # File selection
        self.lbl_file = ctk.CTkLabel(self, text = "File to verify:")
        self.txt_file = ctk.CTkEntry(self, width = 250)
        self.btn_file = ctk.CTkButton(self, text = "Browse…", command = self.__open_file)

        # Output text box
        self.lbl_output = ctk.CTkLabel(self, text = "Result:")
        self.txt_output = ctk.CTkTextbox(self, width = 250, height = 75, state = tk.DISABLED, wrap = "word")
        self.lbl_info = ctk.CTkLabel(self, text = "-")

        # Buttons
        self.btn_verify = ctk.CTkButton(self, text = "Verify", command = self.__verify)

        """
        GUI LAYOUTING
        """
        self.lbl_title.grid(row = 0, column = 0, sticky = tk.W, columnspan = 2, padx = 5, pady = (5, 10))

        self.lbl_file.grid(row = 1, column = 0, sticky = tk.W, padx = 5, pady = 2)
        self.txt_file.grid(row = 1, column = 1, sticky = tk.EW, padx = (5, 2), pady = 2)
        self.btn_file.grid(row = 1, column = 2, sticky = tk.E, padx = (2, 5), pady = 2)

        self.lbl_output.grid(row = 2, column = 0, sticky = tk.NW, padx = 5, pady = 2)
        self.txt_output.grid(row = 2, column = 1, columnspan = 2, sticky = tk.EW, padx = 5, pady = 2)

        self.lbl_info.grid(row = 3, column = 2, sticky = tk.EW, padx = 5, pady = 2)

        self.btn_verify.grid(row = 4, column = 2, sticky = tk.EW, padx = 5, pady = (10, 5))

    def __open_file(self):
        """Opens a file dialog to select a file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select a file to integrity check…", defaultextension = ".pdf", filetypes = [("Portable Document Format", ".pdf")])
        self.txt_file.delete(0, tk.END)
        self.txt_file.insert(0, file_path)

    def __verify(self):
        """Verifies the integrity of the selected file."""
        file_path = self.txt_file.get()
        integrous, output_text = integrity.verify_pdf(file_path)
        
        self.txt_output.configure(state = tk.NORMAL)
        self.txt_output.delete("1.0", tk.END)
        self.txt_output.insert("1.0", output_text)
        self.txt_output.configure(state = tk.DISABLED)

        if integrous:
            self.lbl_info.configure(text = "PASS", fg_color = "green")
        else:
            self.lbl_info.configure(text = "FAIL", fg_color = "red")