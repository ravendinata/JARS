import pandas as pd
from termcolor import colored

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

        try:
            self.course_info = pd.read_excel(self.grader_report_path, sheet_name = "Course Information", index_col = 0, header = 0, nrows = 7)
            self.students = pd.read_excel(self.grader_report_path, sheet_name = "Student List", index_col = 0, header = 0, usecols = "A:C")
            self.data_sna = pd.read_excel(self.grader_report_path, sheet_name = "Skills and Assessment", index_col = 0, header = 0)
            self.data_pd = pd.read_excel(self.grader_report_path, sheet_name = "Personal Development", index_col = 0, header = 0, usecols = "A:I")
            self.data_final_grades = pd.read_excel(self.grader_report_path, sheet_name = "Final Grades", index_col = 0, header = 0, usecols = "A,H:I")
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
        data = self.data_sna.loc[student, assessment]

        if raw:
            return data
        
        if data == "X":
            print(colored(f"Warning: {student} has no grade for {assessment}! Please check the grader report.", "red"))
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
        data = self.data_pd.loc[student, item]

        if raw:
            return data
        
        if data == 0:
            print(colored(
                f"Warning: {student} has no grade for {item}! Please check the grader report. "
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
        data = self.data_final_grades.loc[student, format]

        if raw:
            return data

        if data == 0:
            if format == "Final Score":
                print(colored(
                    f"Warning: {student} has no final score! Please check the grader report. "
                    "Due to this, the final score will be left blank to prevent software errors.", 
                    "red"
                ))
            elif format == "Letter Grade":
                print(colored(
                    f"Warning: {student} has no letter grade! Please check the grader report. "
                    "Due to this, the letter grade will be left blank to X to prevent software errors.", 
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

        print("[  ] Validating data...")
        valid = True
        count = 0

        print(f"Course Info:\n{self.course_info}\n")
        print(f"Student List:\n{self.students}\nStudent Count: {self.count_students()}\n")

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
            if self.get_final_grade(student, "Final Score", raw = True) == 0:
                count += 1
                valid = False
                output_text = f"Warning [{count}]: {student} has no final score! Please check the grader report."
                if callback is not None:
                    callback(output_text)
                print(colored(output_text, "red"))

        # Check if all students have a letter grade
        for student in self.students.index:
            if self.get_final_grade(student, "Letter Grade", raw = True) == 0:
                count += 1
                valid = False
                output_text = f"Warning [{count}]: {student} has no letter grade! Please check the grader report."
                if callback is not None:
                    callback(output_text)
                print(colored(output_text, "red"))

        # Check if all students have a grade for each SNA
        for student in self.students.index:
            for assessment in self.data_sna.columns:
                if self.get_grade_sna(student, assessment, raw = True) == "X":
                    count += 1
                    valid = False
                    output_text = f"Warning [{count}]: {student} has no grade for goal '{assessment}'! Please check the grader report."
                    if callback is not None:
                        callback(output_text)
                    print(colored(output_text, "red"))

        # Check if all students have a grade for each PD item
        for student in self.students.index:
            for item in self.data_pd.columns:
                if self.get_grade_pd(student, item, raw = True) == 0:
                    count += 1
                    valid = False
                    output_text = f"Warning [{count}]: {student} has no grade for personal development item '{item}'! Please check the grader report."
                    if callback is not None:
                        callback(output_text)
                    print(colored(output_text, "red"))

        self.data_valid = valid
        print(colored(f"Validation Pass: {valid}", "red" if not valid else "green"), "\n", colored(f"Warnings: {count}\n", "yellow") if not valid else "")
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

        # Data type conversion
        self.course_info = self.course_info.astype(str)
        self.students = self.students.astype(str)
        self.data_pd = self.data_pd.astype(int)
        self.data_final_grades["Final Score"] = self.data_final_grades["Final Score"].round(0).astype(int)
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