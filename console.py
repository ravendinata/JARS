"""
This program is a console application variant of the JARS program.

It is meant to be used in the command line as an alternative to the GUI application.
The console application is more suitable for batch processing of files by using a batch script.
"""

__version__ = "1.0.0"
__author__ = "Raven Limadinata"

from pathlib import Path

import processor

def get_file_path():
    """
    Prompts the user to enter the source file path and returns it.

    Returns:
        str: The source file path entered by the user.
    """

    file_path = input("\nEnter the source file path.\n(You don't need to remove the quotes (\"...\")): ")
    file_path = file_path.strip('\"')
    
    if not Path(file_path).exists():
        raise ValueError("Invalid file path.")
    
    return file_path

def get_output_file_path(source_file):
    """
    Prompts the user to enter the output file path and returns it.

    Args:
        source_file (str): The path of the source file.

    Returns:
        str: The output file path entered by the user.
    """

    output_file_path = input("\nEnter the output file path (Leave blank to use same name and path as input): ")
    
    if output_file_path == "":
        output_file_path = source_file[:-4] + "_processed.xlsx"
    if not output_file_path.endswith(".xlsx"):
        output_file_path += ".xlsx"
    
    return output_file_path

def check_overwrite(output_file_path):
    """
    Checks if the output file already exists and prompts the user for confirmation to overwrite it.

    Args:
        output_file_path (str): The path of the output file.

    Raises:
        ValueError: If the user chooses not to overwrite the file.
    """

    if Path(output_file_path).exists():
        overwrite = input("\nFile already exists. Do you want to overwrite the file? (y/n): ")
        if overwrite.lower() != "y":
            raise ValueError("File not overwritten.")

def get_mode():
    """
    Prompts the user to enter the report file type and returns the corresponding mode.

    Returns:
        int: The mode representing the report file type.
    """

    mode = int(input("\nWhat report file is this?\n1. Cohort All Assignment Grades\n2. Cohort All Final Grades\nPlease enter appropriate file type: "))
    return mode

def get_adjust_cell_widths():
    """
    Prompts the user to choose whether to auto adjust cell widths and returns the corresponding boolean value.

    Returns:
        bool: True if the user chooses to auto adjust cell widths, False otherwise.
    """

    adjust_cell_widths = input("\nDo you want to auto adjust cell widths? (y/n): ")
    if adjust_cell_widths.lower() == "y":
        return True
    else:
        return False

# Main
source_file = get_file_path()
mode = get_mode()
output_file_path = get_output_file_path(source_file)
check_overwrite(output_file_path)

print(f"\n[  ] Operation started!")
proc = processor.Processor(source_file, output_file_path, mode)
proc.generate_xlsx(adjust_cell_widths = get_adjust_cell_widths())
print(f"[OK] Done. Output file saved at {output_file_path}")