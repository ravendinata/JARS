import openpyxl

import pandas as pd
import customtkinter as ctk
import tkinter as tk

from tkinter import filedialog, messagebox

def process_grouped_data(grouped, output_file_path, value, index, column, sort = False):
    print("[**] Creating output file...")
    with pd.ExcelWriter(output_file_path) as writer:
        for key, item in grouped:
            print(f"[  ] Processing {key}...")
            item = grouped.get_group(key)
            shaped = item.pivot_table(value, index, column, sort = sort)
            shaped.to_excel(writer, sheet_name = key)

    print("[OK] Processing complete!")

def process_file(file_path, output_file_path, mode):
    print("[..] Processing source file...")
    df = pd.read_csv(file_path)

    match mode:
        case 1:
            print("[  ] Grouping data...")
            grouped = df.groupby("course")
            process_grouped_data(grouped, output_file_path, "Grade", ["student name", "class"], "item name")
        
        case 2:
            print("[  ] Grouping data...")
            grouped = df.groupby("class")
            process_grouped_data(grouped, output_file_path, "Grade", "student name", "course")

        case _:
            raise ValueError("[ER] Invalid report file type.")
        
    # Metadata
    print("[  ] Adding metadata...")
    workbook = openpyxl.load_workbook(output_file_path)
    workbook.properties.creator = "JAC Academic Reporting System (JARS)"
    workbook.properties.subject = "JARS Report"
    workbook.properties.keywords = "JARS; Academic Report"
    
    workbook.save(output_file_path)
    print("[OK] Metadata added!")

# UI functions

def browse_file():
    file_path = tk.filedialog.askopenfilename(defaultextension = ".csv", filetypes = [("Comma-Separated Value", "*.csv")])
    txt_source_path.delete(0, tk.END)  # Remove current file path in entry
    txt_source_path.insert(0, file_path)  # Insert the file path into entry

def save_file():
    file_path = tk.filedialog.asksaveasfilename(defaultextension = ".xlsx", filetypes = [("Microsoft Excel Document", "*.xlsx")])
    txt_output_path.delete(0, tk.END)  # Remove current file path in entry
    txt_output_path.insert(0, file_path)  # Insert the file path into entry

def process():
    source_file = txt_source_path.get()
    mode = mode_var.get()
    output_file_path = txt_output_path.get()
    try:
        # check_overwrite(output_file_path)
        process_file(source_file, output_file_path, mode)
        tk.messagebox.showinfo("Success", f"Done. Output file saved at {output_file_path}")
    except Exception as e:
        tk.messagebox.showerror("Error", str(e))

# Root window setup
root = ctk.CTk()
root.title("JARS Report Processor")

root.columnconfigure(0, weight = 1)
root.columnconfigure(1, weight = 2)
root.columnconfigure(2, weight = 1)

# Sourcefile path
lbl_source = ctk.CTkLabel(root, text = "Source file path:")
txt_source_path = ctk.CTkEntry(root, width = 250)
btn_browse_source = ctk.CTkButton(root, text = "Browse...", width = 100, command = browse_file)

# Mode
lbl_mode = ctk.CTkLabel(root, text="Report file type:")
mode_var = tk.IntVar()
rdo_mode_1 = ctk.CTkRadioButton(root, text = "Cohort All Assignment Grades", variable = mode_var, value = 1)
rdo_mode_2 = ctk.CTkRadioButton(root, text = "Cohort All Final Grades", variable = mode_var, value = 2)

# Output file path
lbl_output = ctk.CTkLabel(root, text = "Output file path:")
txt_output_path = ctk.CTkEntry(root, width = 250)
btn_browse_output = ctk.CTkButton(root, text = "Browse...", width = 100, command = save_file)

# Process button
btn_process = ctk.CTkButton(root, text = "Process", width = 100, command = process)

# Layout
lbl_source.grid(row = 0, column = 0, sticky = tk.W, padx = 5, pady = 2)
txt_source_path.grid(row = 0, column = 1, sticky = tk.W, pady = 2)
btn_browse_source.grid(row = 0, column = 2, sticky = tk.EW, padx = 5, pady = 2)

lbl_mode.grid(row = 1, column = 0, sticky = tk.W, padx = 5, pady = 2)
rdo_mode_1.grid(row = 2, column = 0, columnspan = 2, sticky = tk.W, padx = 5, pady = 2)
rdo_mode_2.grid(row = 3, column = 0, columnspan = 2, sticky = tk.W, padx = 5, pady = 2)

lbl_output.grid(row = 4, column = 0, sticky = tk.W, padx = 5, pady = 2)
txt_output_path.grid(row = 4, column = 1, sticky = tk.W, pady = 2)
btn_browse_output.grid(row = 4, column = 2, sticky = tk.EW, padx = 5, pady = 2)

btn_process.grid(row = 5, column = 2, sticky = tk.EW, padx = 5, pady = 10)

root.mainloop()