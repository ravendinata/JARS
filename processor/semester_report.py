import pandas as pd

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT

import config
import processor.helper.document as document_helper
import processor.helper.comment_generator as cgen

class Generator:
    """
    Report generator class for JARS data processor

    This class is responsible for generating a standardized DOCX file from a templated XLSX grader report.
    """

    def __init__(self, grader_report_path, output_path):
        """
        Initialize the generator instance.

        Args:
            grader_report_path (str): The path of the grader report.
            output_path (str): The path of the output file.

        Returns:
            Generator: The initialized generator instance.
        """

        print("[  ] Initializing generator...")

        self.grader_report_path = grader_report_path
        self.output_path = output_path

        self.course_info = pd.read_excel(self.grader_report_path, sheet_name = "Course Information", index_col = 0, header = 0)
        self.students = pd.read_excel(self.grader_report_path, sheet_name = "Student List", index_col = 0, header = 0)
        self.data_sna = pd.read_excel(self.grader_report_path, sheet_name = "Skills and Assessment", index_col = 0, header = 0)
        self.data_pd = pd.read_excel(self.grader_report_path, sheet_name = "Personal Development", index_col = 0, header = 0)
        self.data_final_grades = pd.read_excel(self.grader_report_path, sheet_name = "Final Grades", index_col = 0, header = 0)
        self.data_comment_mapping = pd.read_excel(self.grader_report_path, sheet_name = "Comment Mapping", index_col = 0, header = 0)
        self.__prepare_data()

        print("[OK] Report generator initialized!")

    def get_course_info(self, item):
        return self.course_info.loc[item, "Value"]
    
    def get_student_info(self, student, item):
        return self.students.loc[student, item]
    
    def get_grade_sna(self, student, assessment):
        return self.data_sna.loc[student, assessment]
    
    def get_grade_pd(self, student, item):
        return self.data_pd.loc[student, item]
    
    def get_final_grade(self, student, format):
        return self.data_final_grades.loc[student, format]
    
    def __prepare_data(self):
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

        # Fill NaN values with 0
        self.data_final_grades.fillna(0, inplace = True)

        # Data type conversion
        self.data_final_grades["Final Score"] = self.data_final_grades["Final Score"].round(0).astype(int)

    def generate_all(self, callback = None):
        job_count = len(self.students.index)
        
        for i, student in enumerate(self.students.index):
            status_message = f"Generating report for {student}…"
            callback(i, job_count, status_message)
            print(f"Progress: {i / job_count * 100}%")
            print(status_message)
            self.generate_for_student(student)
        
        callback(job_count, job_count, "Done!")
        print(f"Progress: 100%")

    def generate_for_student(self, student_name):
        document = Document()
        document = document_helper.setup_page(document, 'a4')
        
        section = document.sections[0]

        header_content = section.header.paragraphs[0]
        header_content.alignment = WD_TABLE_ALIGNMENT.CENTER
        logo_run = header_content.add_run()
        logo_run.add_picture(config.get_config("logo_path"), width = Cm(5.56))
        
        # Styles
        style = document.styles["Normal"]
        font = style.font
        font.name = "Calibri"
        font.size = Pt(11)

        top_spacer = document.add_paragraph()
        top_spacer.paragraph_format.space_before = Pt(0)
        top_spacer.paragraph_format.space_after = Pt(0)

        # Course Information Section
        ci_table = document.add_table(rows = 3, cols = 5)
        ci_table.style = "Table Grid"
        ci_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        ci_table.autofit = False
        ci_table.allow_autofit = False
        ci_table_col_widths = [Cm(3), Cm(8), Cm(3), Cm(1.5), Cm(1.5)]

        ci_table.cell(0, 0).text = "Student"
        ci_table.cell(0, 0).paragraphs[0].runs[0].bold = True
        ci_table.cell(0, 1).text = student_name

        ci_table.cell(1, 0).text = "Grade"
        ci_table.cell(1, 0).paragraphs[0].runs[0].bold = True
        ci_table.cell(1, 1).text = self.get_course_info("Grade")
        
        ci_table.cell(2, 0).text = "School Year"
        ci_table.cell(2, 0).paragraphs[0].runs[0].bold = True
        ci_table.cell(2, 1).text = self.get_course_info("School Year")

        ci_table.cell(0, 2).text = "Semester"
        ci_table.cell(0, 2).paragraphs[0].runs[0].bold = True
        ci_table.cell(0, 3).merge(ci_table.cell(0, 4)).text = str(self.get_course_info("Semester"))
        ci_table.cell(0, 3).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

        ci_table.cell(1, 2).text = "Subject"
        ci_table.cell(1, 2).paragraphs[0].runs[0].bold = True
        ci_table.cell(1, 3).merge(ci_table.cell(1, 4)).text = self.get_course_info("Subject")
        ci_table.cell(1, 3).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
        
        ci_table.cell(2, 2).text = "Assessment"
        ci_table.cell(2, 2).paragraphs[0].runs[0].bold = True
        ci_table.cell(2, 3).text = str(self.get_final_grade(student_name, "Final Score"))
        ci_table.cell(2, 3).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
        ci_table.cell(2, 4).text = self.get_final_grade(student_name, "Letter Grade")
        ci_table.cell(2, 4).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

        for i in range(0, 3):
            for j in range(0, 5):
                ci_table.cell(i, j).width = ci_table_col_widths[j]

        # Subject Description Section
        sd_header = document.add_paragraph()
        sd_header.add_run("SUBJECT DESCRIPTION").bold = True
        sd_header.paragraph_format.space_before = Pt(18)
        sd_header.paragraph_format.space_after = Pt(0)
        
        sd_table = document.add_table(rows = 1, cols = 1)
        sd_table.style = "Table Grid"
        sd_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sd_table.autofit = False
        sd_table.allow_autofit = False
        sd_table.rows[0].height = Cm(1.5)
        sd_table.cell(0, 0).text = self.get_course_info("Subject Description")
        sd_table.cell(0, 0).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.LEFT
        sd_table.cell(0, 0).width = Cm(17)

        # Skills and Assessment Section
        sna_header = document.add_paragraph()
        sna_header.add_run("SKILLS AND ASSESSMENT").bold = True
        sna_header.paragraph_format.space_before = Pt(18)
        sna_header.paragraph_format.space_after = Pt(0)
        
        sna_table = document.add_table(rows = len(self.data_sna.columns), cols = 2)
        sna_table.style = "Table Grid"
        sna_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sna_table.autofit = False
        sna_table.allow_autofit = False

        student_sna = {}
        
        for i, assessment in enumerate(self.data_sna.columns):
            sna_table.cell(i, 0).text = assessment
            sna_table.cell(i, 0).width = Cm(15)
            sna_table.cell(i, 1).text = str(self.get_grade_sna(student_name, assessment))
            sna_table.cell(i, 1).width = Cm(2)
            sna_table.cell(i, 1).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
            student_sna[assessment] = self.get_grade_sna(student_name, assessment)

        # Personal Development Section
        pd_header = document.add_paragraph()
        pd_header.add_run("PERSONAL DEVELOPMENT").bold = True
        pd_header.paragraph_format.space_before = Pt(18)
        pd_header.paragraph_format.space_after = Pt(0)
        
        pd_table = document.add_table(rows = len(self.data_pd.columns) + 1, cols = 6)
        pd_table.style = "Table Grid"
        pd_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        pd_table.autofit = False
        pd_table.allow_autofit = False

        pd_table_header = pd_table.rows[0].cells
        pd_table_header[0].text = "Item"
        pd_table_header[1].text = "NI"
        pd_table_header[2].text = "S"
        pd_table_header[3].text = "G"
        pd_table_header[4].text = "VG"
        pd_table_header[5].text = "E"

        for i in range(0, 6):
            pd_table_header[i].paragraphs[0].runs[0].bold = True
            pd_table_header[i].paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
            pd_table_header[i].width = Cm(1) if i != 0 else Cm(12)

        for i, item in enumerate(self.data_pd.columns):
            pd_table.cell(i + 1, 0).text = item
            pd_table.cell(i + 1, 0).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.LEFT
            pd_table.cell(i + 1, 0).width = Cm(12)
            pd_table.cell(i + 1, self.get_grade_pd(student_name, item)).text = "✔"
            pd_table.cell(i + 1, self.get_grade_pd(student_name, item)).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
            
        for i in range(1, len(self.data_pd.columns) + 1):
            for j in range(1, 6):
                pd_table.cell(i, j).width = Cm(1)

        # Teacher's Comments Section
        tc_header = document.add_paragraph()
        tc_header.add_run("TEACHER'S COMMENTS").bold = True
        tc_header.paragraph_format.space_before = Pt(18)
        tc_header.paragraph_format.space_after = Pt(0)

        tc_table = document.add_table(rows = 1, cols = 1)
        tc_table.style = "Table Grid"
        tc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        tc_table.autofit = False
        tc_table.allow_autofit = False
        tc_table.rows[0].height = Cm(2)
        tc_table.cell(0, 0).text = cgen.CommentGenerator(student_name = student_name, 
                                                         short_name = self.get_student_info(student_name, "Short Name"),
                                                         gender = self.get_student_info(student_name, "Gender"),
                                                         comment_mapping = self.data_comment_mapping,
                                                         student_result = student_sna,
                                                         student_final_grade = self.get_final_grade(student_name, "Final Score")
                                                        ).generate_comment()
        tc_table.cell(0, 0).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.LEFT
        tc_table.cell(0, 0).width = Cm(17)

        # Acknowledgement Section
        ak_header = document.add_paragraph()
        ak_header.add_run("ACKNOWLEDGEMENT").bold = True
        ak_header.paragraph_format.space_before = Pt(18)
        ak_header.paragraph_format.space_after = Pt(0)

        ak_table = document.add_table(rows = 2, cols = 3)
        ak_table.style = "Table Grid"
        ak_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        ak_table.autofit = False
        ak_table.allow_autofit = False

        # Adjust row height
        for row in ak_table.rows:
            row.height = Cm(1)

        ak_table.cell(0, 0).paragraphs[0].add_run("Teacher:").bold = True
        ak_table.cell(0, 0).add_paragraph(self.get_course_info("Teacher"))
        ak_table.cell(0, 1).text = "Signature"
        ak_table.cell(0, 2).text = "Date"

        ak_table.cell(1, 0).paragraphs[0].add_run("Parent:").bold = True
        ak_table.cell(1, 1).text = "Signature"
        ak_table.cell(1, 2).text = "Date"

        # Legend Section
        lg_header = document.add_paragraph()
        lg_header.add_run("GRADING SYSTEM").bold = True
        lg_header.paragraph_format.space_before = Pt(18)
        lg_header.paragraph_format.space_after = Pt(0)

        lg_table = document.add_table(rows = 6, cols = 4)
        lg_table.style = "Table Grid"
        lg_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        lg_table.autofit = False
        lg_table.allow_autofit = False

        lg_table.cell(0, 0).merge(lg_table.cell(0, 1)).text = "Skills and Assessment"
        lg_table.cell(0, 0).paragraphs[0].runs[0].bold = True
        lg_table.cell(0, 0).width = Cm(8.5)
        lg_table.cell(0, 2).merge(lg_table.cell(0, 3)).text = "Personal Development"
        lg_table.cell(0, 2).paragraphs[0].runs[0].bold = True
        lg_table.cell(0, 2).width = Cm(8.5)

        lg_table.cell(1, 0).text = "95-100"
        lg_table.cell(1, 1).text = "A"
        lg_table.cell(1, 2).text = "Excellent"
        lg_table.cell(1, 3).text = "E"

        lg_table.cell(2, 0).text = "85-95"
        lg_table.cell(2, 1).text = "B"
        lg_table.cell(2, 2).text = "Very Good"
        lg_table.cell(2, 3).text = "VG"

        lg_table.cell(3, 0).text = "75-85"
        lg_table.cell(3, 1).text = "C"
        lg_table.cell(3, 2).text = "Good"
        lg_table.cell(3, 3).text = "G"

        lg_table.cell(4, 0).text = "65-75"
        lg_table.cell(4, 1).text = "D"
        lg_table.cell(4, 2).text = "Satisfactory"
        lg_table.cell(4, 3).text = "S"

        lg_table.cell(5, 0).text = "Below 65"
        lg_table.cell(5, 1).text = "E"
        lg_table.cell(5, 2).text = "Needs Improvement"
        lg_table.cell(5, 3).text = "NI"

        for i in range(0, 6):
            for j in range(0, 4):
                lg_table.cell(i, j).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
                lg_table.cell(i, j).paragraphs[0].runs[0].font.size = Pt(9)

        document.save(f"{self.output_path}/{student_name}.docx")