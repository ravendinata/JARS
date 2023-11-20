import os
import pandas as pd

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

# Main

source_file = input("Enter the source file path: ")

if (os.path.exists(source_file)):
    print("[OK] Valid file path.")
else:
    print("[ER] Invalid file path.")
    exit()

print("What report file is this?\n1. Cohort All Assignment Grades\n2. Cohort All Final Grades")
mode = int(input("Please enter appropriate file type: "))

output_file_path = input("Enter the output file path: ")

if (output_file_path[-5:] != ".xlsx"):
    output_file_path += ".xlsx"

if (os.path.exists(output_file_path)):
    print("[??] File already exists.")
    overwrite = input("Do you want to overwrite the file? (y/n): ")
    if (overwrite == "y"):
        print("[OK] Overwriting file.")
    elif (overwrite == "n"):
        print("[ER] File not overwritten.")
        exit()
    else:
        raise ValueError("[ER] Invalid input.")

print("[OK] Creating file. Please wait...")
process_file(source_file, output_file_path, mode)
print(f"[OK] File saved to {output_file_path}")