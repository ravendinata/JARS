import openpyxl
import os

import pandas as pd

from pathlib import Path

# Functions

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

def get_file_path():
    file_path = input("\nEnter the source file path.\n(You don't need to remove the quotes (\"...\")): ")
    file_path = file_path.strip('\"')
    
    if not Path(file_path).exists():
        raise ValueError("Invalid file path.")
    
    return file_path

def get_mode():
    mode = int(input("\nWhat report file is this?\n1. Cohort All Assignment Grades\n2. Cohort All Final Grades\nPlease enter appropriate file type: "))
    return mode

def get_output_file_path(source_file):
    output_file_path = input("\nEnter the output file path (Leave blank to use same name and path as input): ")
    
    if output_file_path == "":
        output_file_path = source_file[:-4] + "_processed.xlsx"
    if not output_file_path.endswith(".xlsx"):
        output_file_path += ".xlsx"
    
    return output_file_path

def check_overwrite(output_file_path):
    if Path(output_file_path).exists():
        overwrite = input("\nFile already exists. Do you want to overwrite the file? (y/n): ")
        if overwrite.lower() != "y":
            raise ValueError("File not overwritten.")

# Main

source_file = get_file_path()
mode = get_mode()
output_file_path = get_output_file_path(source_file)
check_overwrite(output_file_path)

print(f"\n[  ] Operation started!")
process_file(source_file, output_file_path, mode)
print(f"[OK] Done. Output file saved at {output_file_path}")