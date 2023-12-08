"""
This module is a console handler for the JARS program report generator.

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
import processor.grader_report as grader_report
import processor.semester_report as processor
import processor.helper.language_tool_master as ltm

help_text = """
HELP PAGE
=========
This script generates a subject report from the grader report file and saves the results to the given output folder.

=========
USAGE
=========
Format:
report_generator.py -s <source_file_path> -o <output_file_path> -a -all --student <student_name> --help

Options:
-h, --help
    Displays this help page.
-s, --source <source_file_path>
    Specifies the path of the source file (grader report). Must be an Excel file (.xlsm or .xlsx).
    Note: The source file must contain a sheet named 'Comment Mapping'.
    Example:
        C:/Users/John Doe/Desktop/Grader Report P1A Art Sample.xlsm
        or
        C:/Users/John Doe/Desktop/Grader Report P1A Art Sample.xlsx
-o, --output <output_file_path>
    Specifies the path of the output file (Excel file). Must include the file name and extension.
    Note: If the file already exists, it will be overwritten.
    Example:
        C:/Users/John Doe/Desktop/Test Result.xlsx
-a, --autocorrect
    Specifies whether to autocorrect the comments.
--all
    Specifies whether to generate reports for all students in the grader report.
--student <student_name>
    Specifies the name of the student to generate the report for.
    Note: This option is only available if the --all option is not specified.
    Example:
        --student "John Doe"
-f, --force
    Specifies whether to force the program to proceed with the operation even if invalid data is detected in the source file.

Example:
report_generator.py -s "C:/Grader Report P1A Art Sample.xlsm" -o "C:/Reports" -a --student "John Doe" --force
or
console.py -t report_generator -s "C:/Grader Report.xlsm" -o "C:/Reports" -a --all
"""

def get_mode():
    """
    Prompts user whether to generate one report for a student or for all students inside the grader report.

    Returns:
        int: The mode representing the report file type.
    """

    mode = int(input("\nGenerate for:\n1. All students in grader report\n2. Single student\nPlease enter appropriate file type: "))
    return mode

def prompt_autocorrect():
    """
    Prompts user whether to autocorrect the comments.

    Returns:
        bool: True if the user chooses to autocorrect the comments, False otherwise.
    """

    autocorrect = input("\nDo you want to autocorrect the comments? (y/n): ")
    if autocorrect.lower() == "y":
        return True
    else:
        return False

def run():
    source_file = con.get_source_file_path()
    mode = get_mode()
    output_file_path = input("\nEnter the path to the output folder: ")

    if mode == 2:
        student_name = input("\nEnter the student name to generate the report for: ")

    autocorrect = prompt_autocorrect()

    print(f"\n[  ] Operation started!")
    start_time = time.time()

    gr = grader_report.GraderReport(source_file)
    proc = processor.Generator(output_file_path, gr)
    if mode == 1:
        proc.generate_all(autocorrect = autocorrect)
    elif mode == 2:
        proc.generate_for_student(student_name = student_name, autocorrect = autocorrect)

    print(f"[OK] Operation completed in {time.time() - start_time} seconds.")
    print(f"[OK] Done. Output file saved at {output_file_path}")

    ltm.close_tool()

short_args = "hs:o:afp"
long_args = ["help", "source=", "output=", "autocorrect", "all", "student=", "force", "pdf"]

def main(argv):
    source_file_path = ""
    output_file_path = ""
    autocorrect = False
    generate_all = False
    force = False
    pdf = False

    print(f"Argument List: {argv}")

    try:
        opts, args = getopt.getopt(argv, short_args, long_args)
    except getopt.GetoptError:
        print("Error! Invalid argument(s).")
        print("report_generator.py -s <source_file_path> -o <output_file_path> -a -all --student --help")
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(help_text)
            sys.exit()
        elif opt in ("-s", "--source"):
            source_file_path = arg
        elif opt in ("-o", "--output"):
            output_file_path = arg
        elif opt in ("-a", "--autocorrect"):
            autocorrect = True
        elif opt in ("--all"):
            generate_all = True
        elif opt in ("--student"):
            student_name = arg
        elif opt in ("-f", "--force"):
            force = True
        elif opt in ("-p", "--pdf"):
            pdf = True

    if source_file_path == "" or output_file_path == "":
        print("Error! No output and/or source file path specified.")
        print("report_generator.py -s <source_file_path> -o <output_file_path>")
        sys.exit(2)

    gr = grader_report.GraderReport(source_file_path)
    proc = processor.Generator(output_file_path, gr)
    if generate_all:
        proc.generate_all(autocorrect = autocorrect, force = force, convert_to_pdf = pdf)
    else:
        proc.generate_for_student(student_name = student_name, autocorrect = autocorrect, force = force, convert_to_pdf = pdf)

    ltm.close_tool()