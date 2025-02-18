import json

import pandas as pd
from termcolor import colored
from openpyxl import load_workbook

class GraderReport:
    def __init__(self, grader_report_path, skip_validation = False, callback = None):
        """
        Initialize the grader report instance.

        Args:
            grader_report_path (str): The path of the grader report.
            output_path (str): The path of the output file.

        Returns:
            Generator: The initialized grader report instance.
        """

        print("[  ] Initializing generator...")

        self.grader_report_path = grader_report_path
        self.__data_broken = False

        # Get the file version from the status property
        props = load_workbook(self.grader_report_path).properties
        self._version = props.version if props.version is not None else "1.0"

        if self._version != "1.2":
            print(colored(f"Warning: Grader report version is {self._version}. Please make sure to use the latest version of the grader report template.", "yellow"))

        try:
            pd.options.display.float_format = '{:.2f}'.format

            self.course_info = pd.read_excel(self.grader_report_path, sheet_name = "Course Information", index_col = 0, header = 0, nrows = 7)
            self.students = pd.read_excel(self.grader_report_path, sheet_name = "Student List", index_col = 0, header = 0, usecols = "A:C")
            self.data_sna = pd.read_excel(self.grader_report_path, sheet_name = "Skills and Assessment", index_col = 0, header = 0)
            self.data_pd = pd.read_excel(self.grader_report_path, sheet_name = "Personal Development", index_col = 0, header = 0, usecols = "A:I")
            self.data_final_grades = pd.read_excel(self.grader_report_path, sheet_name = "Final Grades", index_col = 0, header = 0, usecols = "A,H:I", dtype = {"Final Score": float, "Letter Grade": str})
            self.data_comment_mapping = pd.read_excel(self.grader_report_path, sheet_name = "Comment Mapping", index_col = 0, header = 0)
        except Exception as e:
            self.__data_broken = True
            output_text = "Error: Unable to read the grader report! Please check that you are using the base template designed for JARS."
            print(colored(output_text, "white", "on_red"))
            print(colored(f"Details: {e}", "red"))
            if callback is not None:
                callback(output_text)
            return
        
        self.__prepare_data()

        if not skip_validation:
            self.validate()
            if not self.data_valid:
                print(colored("\n"
                    f"Attention: There are missing values in the grader report! Please check the grader report again and make sure all data is filled.",
                    "white", "on_red"), "\n")

        print("[OK] Report generator initialized!")

    # Getters
    def get_course_info(self, item):
        """
        Get course information.

        Args:
            item (str): The course information item to get.

        Returns:
            str: The value of the course information item.
        """
        return self.course_info.loc[item, "Value"]
    
    def get_student_info(self, student, item):
        """
        Get student information.
        
        Args:
            student (str): The student to get. Must be the student's full name.
            item (str): The student information item to get.
            
        Returns:
            str: The value of the student information item.
        """
        return self.students.loc[student, item]
    
    def get_grade_sna(self, student, assessment, raw = False):
        """
        Get student grade for SNA. This method will return an empty string if the student has no grade for the SNA
        to prevent software errors when force generating the report.
        
        Args:
            student (str): The student to get. Must be the student's full name.
            assessment (str): The assessment to get.
            raw (bool): Whether to return the raw value or not.
            
        Returns:
            str: The value of the student grade for the SNA.
        """
        try:
            data = self.data_sna.loc[student, assessment]
        except KeyError:
            print(colored(f"Error: {student} not found in the Skills and Assessment sheet! Please check that the student is included in the sheet.", "red"))
            return "SMFS"
        except Exception as e:
            print(colored(f"Error: An error occurred while getting grade for {student}! Details: {e}", "red"))
            return ""

        if raw:
            return data
        
        if data == "X":
            print(colored(f"Error: {student} has no grade for {assessment}! Please check the grader report.", "red"))
            return ""
        else:
            return data
    
    def get_grade_pd(self, student, item, raw = False):
        """
        Get student grade for PD. This method will return 1 if the student has no grade for the PD item
        to prevent software errors when force generating the report.

        Args:
            student (str): The student to get. Must be the student's full name.
            item (str): The PD item to get.
            raw (bool): Whether to return the raw value or not.

        Returns:
            str: The value of the student grade for the PD item.
        """
        try:
            data = self.data_pd.loc[student, item]
        except KeyError:
            print(colored(f"Error: {student} not found in the Personal Development sheet! Please check that the student is included in the sheet.", "red"))
            return "SMFS"
        except Exception as e:
            print(colored(f"Error: An error occurred while getting grade for {student}! Details: {e}", "red"))
            return 0

        if raw:
            return data
        
        if data == 0:
            print(colored(
                f"Error: {student} has no grade for {item}! Please check the grader report. "
                "Due to this, the grade will be set to NI to prevent software errors.", 
                "red"))
            return 1
        else:
            return data
    
    def get_final_grade(self, student, format, raw = False):
        """
        Get student final grade. This method will return 0 if the student has no final grade
        to prevent software errors when force generating the report.
        
        Args:
            student (str): The student to get. Must be the student's full name.
            format (str): The format to get. Can be either "Final Score" or "Letter Grade".
            raw (bool): Whether to return the raw value or not.
        
        Returns:
            str: The value of the student final grade.
        """
        try:
            data = self.data_final_grades.loc[student, format]
        except KeyError:
            print(colored(f"Error: {student} not found in the Final Grades sheet! Please check that the student is included in the sheet.", "red"))
            return "SMFS"
        except Exception as e:
            print(colored(f"Error: An error occurred while getting final grade for {student}! Details: {e}", "red"))
            return 0

        if raw:
            return data

        if data == 0:
            print(colored(
                f"Error: {student} has no {format}! Please check the grader report. "
                "Due to this, the final score will be left blank to prevent software errors.", 
                "red"
            ))
            data = str(" ")
        
        return data
    
    def count_sna(self):
        """
        Get the length of the SNA data.
        
        Returns:
            int: The length of the SNA data.
        """
        return len(self.data_sna.columns)
    
    def count_students(self):
        """
        Get the length of the student data.
        
        Returns:
            int: The length of the student data.
        """
        return len(self.students.index)
    
    # Misc functions
    def validate(self, callback = None):
        """
        Validate the data in the grader report. This method will print warnings if there are missing values in the grader report.
        
        Returns:
            bool: Whether the data is valid or not.
        """
        if self.__data_broken:
            output_text = "Error: Unable to validate the grader report due to data corruption or template incompliance. Please make sure to use the base template designed for JARS."
            print(colored(output_text, "white", "on_red"))
            if callback is not None:
                callback(output_text)
            return False

        print("[  ] Validating data...\n")
        valid = True
        count = 0

        print(f"Grader Report Version: {self._version}\n")
        print(f"Course Info:\n{self.course_info}\n")
        print(f"Student List:\n{self.students}\n\nStudent Count: {self.count_students()}\n")

        if self.count_students() == 0:
            count += 1
            valid = False
            output_text = f"Warning [{count}]: No students found in the grader report! Please check 'Student List' sheet on the grader report. You MUST fill the Student Name, Short Name, and Gender columns."
            if callback is not None:
                callback(output_text)
            print(colored(output_text, "red"))

        # Check if all course information is filled
        for item in self.course_info.index:
            if str(self.get_course_info(item)) == "": # Check if the value is empty
                count += 1
                valid = False
                output_text = f"Warning [{count}]: Course information '{item}' is not filled! Please check the grader report."
                if callback is not None:
                    callback(output_text)
                print(colored(output_text, "red"))

        # Check if all students have a final grade
        for student in self.students.index:
            final_grade_result = self.get_final_grade(student, "Final Score", raw = True)
            if final_grade_result == 0:
                count += 1
                valid = False
                output_text = f"Warning [{count}]: {student} has no final score! Please check the grader report."
                if callback is not None:
                    callback(output_text)
                print(colored(output_text, "red"))
            elif final_grade_result == "SMFS":
                count += 1
                valid = False
                output_text = f"Warning [{count}]: {student} is not found in the 'Final Grades' sheet! Please check that the student is included in the sheet."
                if callback is not None:
                    callback(output_text)
                break

        # Check if all students have a letter grade
        for student in self.students.index:
            letter_grade_result = self.get_final_grade(student, "Letter Grade", raw = True)
            if letter_grade_result == 0:
                count += 1
                valid = False
                output_text = f"Warning [{count}]: {student} has no letter grade! Please check the grader report."
                if callback is not None:
                    callback(output_text)
                print(colored(output_text, "red"))

        # Check if all students have a grade for each SNA
        for student in self.students.index:
            for assessment in self.data_sna.columns:
                sna_result = self.get_grade_sna(student, assessment, raw = True)
                if sna_result == "X":
                    count += 1
                    valid = False
                    output_text = f"Warning [{count}]: {student} has no grade for goal '{assessment}'! Please check the grader report."
                    if callback is not None:
                        callback(output_text)
                    print(colored(output_text, "red"))
                elif sna_result == "SMFS":
                    count += 1
                    valid = False
                    output_text = f"Warning [{count}]: {student} is not found in the 'Skills and Assessment' sheet! Please check that the student is included in the sheet."
                    if callback is not None:
                        callback(output_text)
                    break

        # Check if all students have a grade for each PD item
        for student in self.students.index:
            for item in self.data_pd.columns:
                pd_result = self.get_grade_pd(student, item, raw = True)
                if pd_result == 0:
                    count += 1
                    valid = False
                    output_text = f"Warning [{count}]: {student} has no grade for personal development item '{item}'! Please check the grader report."
                    if callback is not None:
                        callback(output_text)
                    print(colored(output_text, "red"))
                elif pd_result == "SMFS":
                    count += 1
                    valid = False
                    output_text = f"Warning [{count}]: {student} is not found in the 'Personal Development' sheet! Please check that the student is included in the sheet."
                    if callback is not None:
                        callback(output_text)
                    break

        for student in self.students.index:
            if not self.__check_sna_compliance(student):
                count += 1
                valid = False
                output_text = f"Warning [{count}]: {student} does not meet the SNA compliance rules! Please check the grader report."
                if callback is not None:
                    callback(output_text)
                print(colored(output_text, "red"))

        self.data_valid = valid
        print(colored(f"Validation Pass: {valid}", "red" if not valid else "green"), "\n", colored(f"Warnings: {count}\n", "yellow") if not valid else "")
        callback(f"\nYou have {count} issues in your grader report. Please check the warnings above." if not valid else "")
        return valid
    
    # Private methods
    def __prepare_data(self):
        """
        Prepare the data for the generator. This method will remove NaN values and strip whitespace from the data.
        """
        # Remove NaN rows
        self.students.dropna(inplace = True)

        # Fill NaN values with default values
        self.course_info.fillna("", inplace = True)
        self.data_final_grades.fillna(0, inplace = True)
        self.data_pd.fillna(0, inplace = True)

        def excel_rounding(x):
            remainder = x % 1
            sub = x - remainder
            if remainder >= 0.5:
                return int(sub + 1)
            else:
                return int(sub)

        # Data type conversion
        self.course_info = self.course_info.astype(str)
        self.students = self.students.astype(str)
        self.data_pd = self.data_pd.astype(int)
        self.data_final_grades["Final Score"] = self.data_final_grades["Final Score"].apply(excel_rounding)
        self.data_comment_mapping = self.data_comment_mapping.astype(str)
        
        # Convert SNA data to string and fill NaN values with "X"
        self.data_sna = self.data_sna.astype(str)
        self.data_sna.fillna("X", inplace = True)

        # Strip whitespace from index
        self.course_info.index = self.course_info.index.str.strip()
        self.students.index = self.students.index.str.strip()
        self.data_sna.index = self.data_sna.index.str.strip()
        self.data_pd.index = self.data_pd.index.str.strip()
        self.data_final_grades.index = self.data_final_grades.index.str.strip()
        self.data_comment_mapping.index = self.data_comment_mapping.index.str.strip()

        # Strip whitespace from columns
        self.course_info.columns = self.course_info.columns.str.strip()
        self.students.columns = self.students.columns.str.strip()
        self.data_sna.columns = self.data_sna.columns.str.strip()
        self.data_pd.columns = self.data_pd.columns.str.strip()
        self.data_final_grades.columns = self.data_final_grades.columns.str.strip()
        self.data_comment_mapping.columns = self.data_comment_mapping.columns.str.strip()

        # Remove unnecessary columns
        unwanted_columns = ["Normalized Grade", "Student Final Grade", "Sanity Check", "Add item…"]
        unwanted_columns += [column for column in self.data_sna.columns if column.startswith("Unnamed")]
        self.data_sna = self.data_sna.drop(columns = unwanted_columns)
                        
    def _load_sna_rules(self):
        """Load SNA compliance rules from JSON file."""
        rules_path = 'resources/sna_rules.json'
        try:
            with open(rules_path, 'r') as f:
                return json.load(f)['rules']
        except FileNotFoundError:
            raise FileNotFoundError(f"Rules file not found at {rules_path}")
        except json.JSONDecodeError:
            raise ValueError("Invalid JSON format in rules file")
        
    def __check_sna_compliance(self, student):
        """
        Check if the SNA data is compliant with the scoring guide.
        
        Args:
            student: The student whose SNA compliance needs to be checked
            
        Returns:
            bool: True if compliant, False otherwise
        """
        rules = self._load_sna_rules()
        sna_count = str(self.count_sna())  # Convert to string to match JSON keys
        student_final_grade = self.get_final_grade(student, "Final Score", raw = False)

        if type(student_final_grade) == str:
            print(colored("Skipping SNA compliance check due to missing final grade", "yellow"))
            return False
        
        grade_counts = {'A': 0, 'B': 0, 'C': 0}
        for assessment in self.data_sna.columns:
            grade = self.get_grade_sna(student, assessment, raw = True)
            if grade in grade_counts:
                grade_counts[grade] += 1

        if sna_count not in rules:
            print(colored(f"Skipping SNA compliance check due to missing rules for {sna_count} SNAs", "yellow"))
            return False
        
        if student_final_grade == 'SMFS':
            print(colored("Skipping SNA compliance check because student is not found in the grader report", "yellow"))
            return False
        
        # Find matching grade range and check requirements
        for grade_range, grade_reqs in rules[sna_count].items():
            min_grade, max_grade = map(int, grade_range.split('-'))
            if min_grade <= student_final_grade < max_grade:
                # Check if all grade counts match requirements
                for grade, req in grade_reqs.items():
                    if type(req) == str:
                        min_count, max_count = map(int, req.split('-'))
                    else:
                        min_count = max_count = int(req)
                    actual_count = grade_counts[grade]
                    if not (min_count <= actual_count <= max_count):
                        return False
                return True
        
        return False