"""
This module is a console handler for the JARS program report formatter.

It is meant to be used in the command line as an alternative to the GUI application.
The console application is more suitable for batch processing of files by using a batch script.
"""

__version__ = "1.0.0"
__author__ = "Raven Limadinata"

if __name__ == "__main__":
    print("This script is not meant to be run directly. Please run this script from console.py file.")
    exit()

import time
import getopt
import sys

import console.helper as con
import processor.formatter as processor

help_text = """
HELP PAGE
=========
This script formats the given report file and saves the results to the given output file.

=========
USAGE
=========
Format:
report_formatter.py -s <source_file_path> -o <output_file_path>

Options:
-h, --help
    Displays this help page.
-s, --source <source_file_path>
    Specifies the path of the source file (report file). Must be an CSV file (.csv).
    Example:
        C:/Users/John Doe/Desktop/Report.csv
-o, --output <output_file_path>
    Specifies the path of the output file (Excel file). Must include the file name and extension.
    Example:
        C:/Users/John Doe/Desktop/Report.xlsx
-p, --profile <report_profile_path>
    Specifies the path of the report profile (JSON file). Must include the file name and extension.
-w, --adjust-widths
    Specifies whether to auto adjust cell widths.

Example:
report_formatter.py -s C:/Users/John Doe/Desktop/Report.csv -o C:/Users/John Doe/Desktop/Formatted Report.xlsx -p C:/Users/John Doe/Desktop/Report Profile.json -w
"""

def get_report_profile():
    """
    Prompts the user to choose a report profile name.

    Returns:
        str: The report profile name
    """

    report_profile = input("Enter the report profile name: ")
    report_profile = report_profile.strip('\"')

    return report_profile

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

def run():
    source_file = con.get_source_file_path()
    profile = get_report_profile()
    output_file_path = con.get_output_file_path(source_file, ".xlsx")
    con.prompt_overwrite(output_file_path)

    adjust_cell_widths = get_adjust_cell_widths()

    print(f"\n[  ] Operation started!")
    start_time = time.time()
    proc = processor.Formatter(source_file, output_file_path, profile)
    proc.generate_xlsx(adjust_cell_widths = adjust_cell_widths)
    print(f"[OK] Operation completed in {time.time() - start_time} seconds.")
    print(f"[OK] Done. Output file saved at {output_file_path}")

short_args = "hs:o:p:w"
long_args = ["help", "source=", "output=", "profile=", "adjust-widths="]

def main(argv):
    source_file_path = ""
    output_file_path = ""
    profile = ""
    adjust_cell_widths = False

    try:
        opts, args = getopt.getopt(argv, short_args, long_args)
    except getopt.GetoptError:
        print("Error! Invalid argument(s).")
        print("report_formatter.py -s <source_file_path> -o <output_file_path> -p <report_profile_path> -w --help")
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(help_text)
            sys.exit()
        elif opt in ("-s", "--source"):
            source_file_path = arg
        elif opt in ("-o", "--output"):
            output_file_path = arg
        elif opt in ("-p", "--profile"):
            profile = arg
        elif opt in ("-w", "--adjust-widths"):
            adjust_cell_widths = True

    if source_file_path == "" or output_file_path == "" or profile == "":
        print("Error! Incomplete requirements.")
        print("report_formatter.py -s <source_file_path> -o <output_file_path> -p <report_profile_path>")
        sys.exit(2)

    proc = processor.Formatter(source_file_path, output_file_path, profile)
    proc.generate_xlsx(adjust_cell_widths = adjust_cell_widths)