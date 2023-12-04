"""
This program is a console application variant of the JARS program.

It is meant to be used in the command line as an alternative to the GUI application.
The console application is more suitable for batch processing of files by using a batch script.
"""

__version__ = "1.0.0"
__author__ = "Raven Limadinata"

import getopt
import sys

import console.report_formatter as report_formatter
import console.report_generator as report_generator

help_text = """
HELP PAGE
=========
This script is a console application for the JARS program.

=========
USAGE
=========
Format:
main.py -i --help
main.py -i --help
main.py -t <tool_name> --help

Options:
-h, --help
    Displays this help page.
-i, --interactive
    Starts the interactive mode.
-t, --tool <tool_name>
    Specifies the tool to run.
    Example:
        --tool report_formatter
        or
        --tool report_generator
Tip: Use the -h or --help option to display the help page for the specified tool.
Note: Pass the arguments for the specified tool after the tool name.

Example:
console.py -t report_generator -s "C:/Grader Report.xlsm" -o "C:/Reports" -a --all
"""

def interactive():
    print("JARS Report Processor\nJAC Academic Reporting System | Version 1.0.0")
    print("\nSelect a tool to open:\n1. Report Formatter\n2. Report Generator")
    tool = int(input("Please enter appropriate tool number: "))

    if tool == 1:
        report_formatter.run()
    elif tool == 2:
        report_generator.run()

    input("\nPress Enter to exit…")

short_args = "hit:"
long_args = ["help", "interactive", "tool="]

def main(argv):
    try:
        opts, args = getopt.getopt(argv, short_args + report_formatter.short_args + report_generator.short_args, 
                                   long_args + report_formatter.long_args + report_generator.long_args)
    except getopt.GetoptError:
        print("Error! Invalid argument(s).")
        print("main.py -t <tool_name> -i --help")
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(help_text)
            sys.exit()
        elif opt in ("-i", "--interactive"):
            interactive()
        elif opt in ("-t", "--tool"):
            if arg == "report_formatter":
                report_formatter.main(sys.argv[3:])
            elif arg == "report_generator":
                report_generator.main(sys.argv[3:])

if __name__ == "__main__":
    main(sys.argv[1:])