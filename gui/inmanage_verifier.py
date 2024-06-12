import os

import customtkinter as ctk
import tkinter as tk

import components.common.integrity as integrity
import components.common.metadata as metadata

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

        # File signature information
        self.lbl_signature_info = ctk.CTkLabel(self, text = "Signature Information:")
        self.lbl_signer = ctk.CTkLabel(self, text = "Signer:")
        self.txt_signer = ctk.CTkEntry(self, width = 250, state = tk.DISABLED)
        self.lbl_signed_on = ctk.CTkLabel(self, text = "Signed On:")
        self.txt_signed_on = ctk.CTkEntry(self, width = 250, state = tk.DISABLED)
        self.lbl_hash = ctk.CTkLabel(self, text = "Hash:")
        self.txt_hash = ctk.CTkEntry(self, width = 250, state = tk.DISABLED)
        self.lbl_serial_number = ctk.CTkLabel(self, text = "Serial Number:")
        self.txt_serial_number = ctk.CTkEntry(self, width = 250, state = tk.DISABLED)

        # Output text box
        self.lbl_output = ctk.CTkLabel(self, text = "Result:")
        self.txt_output = ctk.CTkTextbox(self, width = 250, height = 75, state = tk.DISABLED, wrap = "word")
        self.lbl_info = ctk.CTkLabel(self, text = "-", text_color = "white")

        # Buttons
        self.btn_verify = ctk.CTkButton(self, text = "Verify", command = self.__verify)

        """
        GUI LAYOUTING
        """
        self.lbl_title.grid(row = 0, column = 0, sticky = tk.W, columnspan = 2, padx = 5, pady = (5, 10))

        self.lbl_file.grid(row = 1, column = 0, sticky = tk.W, padx = 5, pady = 2)
        self.txt_file.grid(row = 1, column = 1, sticky = tk.EW, padx = (5, 2), pady = 2)
        self.btn_file.grid(row = 1, column = 2, sticky = tk.E, padx = (2, 5), pady = 2)

        self.lbl_signature_info.grid(row = 2, column = 0, sticky = tk.W, padx = 5, pady = 2, columnspan = 2)
        
        self.lbl_signer.grid(row = 3, column = 0, sticky = tk.W, padx = 5, pady = 2)
        self.txt_signer.grid(row = 3, column = 1, columnspan = 2, sticky = tk.EW, padx = 5, pady = 2)
        
        self.lbl_signed_on.grid(row = 4, column = 0, sticky = tk.W, padx = 5, pady = 2)
        self.txt_signed_on.grid(row = 4, column = 1, columnspan = 2, sticky = tk.EW, padx = 5, pady = 2)
        
        self.lbl_hash.grid(row = 5, column = 0, sticky = tk.W, padx = 5, pady = 2)
        self.txt_hash.grid(row = 5, column = 1, columnspan = 2, sticky = tk.EW, padx = 5, pady = 2)

        self.lbl_serial_number.grid(row = 6, column = 0, sticky = tk.W, padx = 5, pady = 2)
        self.txt_serial_number.grid(row = 6, column = 1, columnspan = 2, sticky = tk.EW, padx = 5, pady = 2)

        self.lbl_output.grid(row = 7, column = 0, sticky = tk.NW, padx = 5, pady = 2)
        self.txt_output.grid(row = 7, column = 1, columnspan = 2, sticky = tk.EW, padx = 5, pady = 2)

        self.lbl_info.grid(row = 8, column = 2, sticky = tk.EW, padx = 5, pady = 2)

        self.btn_verify.grid(row = 9, column = 2, sticky = tk.EW, padx = 5, pady = (10, 5))

    def __open_file(self):
        """Opens a file dialog to select a file."""
        file_path = ctk.filedialog.askopenfilename(title = "Select a file to integrity check…", defaultextension = ".pdf", filetypes = [("Portable Document Format", ".pdf")])
        self.txt_file.delete(0, tk.END)
        self.txt_file.insert(0, file_path)

        if os.path.isfile(file_path):
            # Reset results field
            self.txt_output.configure(state = tk.NORMAL)
            self.txt_output.delete("1.0", tk.END)
            self.txt_output.configure(state = tk.DISABLED)
            self.lbl_info.configure(text = "PENDING", fg_color = "blue")

            # Get metadata
            data = metadata.get_pdf_metadata(file_path)

            # Show signature metadata
            self.txt_signer.configure(state = tk.NORMAL)
            self.txt_signer.delete(0, tk.END)
            self.txt_signer.insert(0, data.get("/Signed By", "N/A"))
            self.txt_signer.configure(state = tk.DISABLED)

            self.txt_signed_on.configure(state = tk.NORMAL)
            self.txt_signed_on.delete(0, tk.END)
            self.txt_signed_on.insert(0, data.get("/Signed On", "N/A"))
            self.txt_signed_on.configure(state = tk.DISABLED)
            
            self.txt_hash.configure(state = tk.NORMAL)
            self.txt_hash.delete(0, tk.END)
            self.txt_hash.insert(0, data.get("/Hash", "N/A"))
            self.txt_hash.configure(state = tk.DISABLED)

            self.txt_serial_number.configure(state = tk.NORMAL)
            self.txt_serial_number.delete(0, tk.END)
            self.txt_serial_number.insert(0, data.get("/Serial Number", "N/A"))
            self.txt_serial_number.configure(state = tk.DISABLED)

    def __verify(self):
        """Verifies the integrity of the selected file."""
        file_path = self.txt_file.get()

        if not os.path.isfile(file_path):
            self.txt_output.configure(state = tk.NORMAL)
            self.txt_output.delete("1.0", tk.END)
            self.txt_output.insert("1.0", "The file does not exist!\nPlease check the file path and select a valid file to verify.")
            self.txt_output.configure(state = tk.DISABLED)
            self.lbl_info.configure(text = "FAIL", fg_color = "red")
            return

        integrous, output_text = integrity.verify_pdf(file_path)
        
        self.txt_output.configure(state = tk.NORMAL)
        self.txt_output.delete("1.0", tk.END)
        self.txt_output.insert("1.0", output_text)
        self.txt_output.configure(state = tk.DISABLED)

        if integrous:
            self.lbl_info.configure(text = "PASS", fg_color = "green")
        else:
            self.lbl_info.configure(text = "FAIL", fg_color = "red")