import pandas as pd

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT

import processor.helper.document as document_helper

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
    
    def generate_all(self):
        for student in self.students.index:
            self.generate_for_student(student)
    
    def __prepare_data(self):
        # Strip whitespace from index
        self.course_info.index = self.course_info.index.str.strip()
        self.students.index = self.students.index.str.strip()
        self.data_sna.index = self.data_sna.index.str.strip()
        self.data_pd.index = self.data_pd.index.str.strip()

        # Strip whitespace from columns
        self.course_info.columns = self.course_info.columns.str.strip()
        self.students.columns = self.students.columns.str.strip()
        self.data_sna.columns = self.data_sna.columns.str.strip()
        self.data_pd.columns = self.data_pd.columns.str.strip()

    def generate_for_student(self, student_name):
        document = Document()
        document = document_helper.setup_page(document, 'a4')
        
        section = document.sections[0]

        header_content = section.header.paragraphs[0]
        header_content.alignment = WD_TABLE_ALIGNMENT.CENTER
        logo_run = header_content.add_run()
        logo_run.add_picture("Z:/Work/JAC/logo.png", width = Cm(5.56))
        
        # Styles
        style = document.styles["Normal"]
        font = style.font
        font.name = "Calibri"
        font.size = Pt(11)

        document.add_paragraph()

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
        ci_table.cell(2, 3).text = "100"
        ci_table.cell(2, 3).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER
        ci_table.cell(2, 4).text = "A"
        ci_table.cell(2, 4).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

        for i in range(0, 3):
            for j in range(0, 5):
                ci_table.cell(i, j).width = ci_table_col_widths[j]

        # Subject Description Section
        sd_header = document.add_paragraph()
        sd_header.add_run("SUBJECT DESCRIPTION").bold = True
        sd_header.paragraph_format.space_before = Pt(24)
        sd_header.paragraph_format.space_after = Pt(0)
        
        sd_table = document.add_table(rows = 1, cols = 1)
        sd_table.style = "Table Grid"
        sd_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sd_table.autofit = False
        sd_table.allow_autofit = False
        sd_table.cell(0, 0).text = self.get_course_info("Subject Description")
        sd_table.cell(0, 0).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.LEFT
        sd_table.cell(0, 0).width = Cm(17)

        # Skills and Assessment Section
        sna_header = document.add_paragraph()
        sna_header.add_run("SKILLS AND ASSESSMENT").bold = True
        sna_header.paragraph_format.space_before = Pt(24)
        sna_header.paragraph_format.space_after = Pt(0)
        
        sna_table = document.add_table(rows = len(self.data_sna.columns), cols = 2)
        sna_table.style = "Table Grid"
        sna_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sna_table.autofit = False
        sna_table.allow_autofit = False
        
        for i, assessment in enumerate(self.data_sna.columns):
            sna_table.cell(i, 0).text = assessment
            sna_table.cell(i, 0).width = Cm(15)
            sna_table.cell(i, 1).text = str(self.get_grade_sna(student_name, assessment))
            sna_table.cell(i, 1).width = Cm(2)
            sna_table.cell(i, 1).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.CENTER

        # Personal Development Section
        pd_header = document.add_paragraph()
        pd_header.add_run("PERSONAL DEVELOPMENT").bold = True
        pd_header.paragraph_format.space_before = Pt(24)
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
        tc_header.paragraph_format.space_before = Pt(24)
        tc_header.paragraph_format.space_after = Pt(0)

        tc_table = document.add_table(rows = 1, cols = 1)
        tc_table.style = "Table Grid"
        tc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        tc_table.autofit = False
        tc_table.allow_autofit = False
        tc_table.cell(0, 0).text = "Student is a hardworking and diligent student. She is always willing to help her peers and is a good role model for her classmates."
        tc_table.cell(0, 0).paragraphs[0].alignment = WD_TABLE_ALIGNMENT.LEFT
        tc_table.cell(0, 0).width = Cm(17)

        # Acknowledgement Section
        ak_header = document.add_paragraph()
        ak_header.add_run("ACKNOWLEDGEMENT").bold = True
        ak_header.paragraph_format.space_before = Pt(24)
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

        document.save(f"{self.output_path}/{student_name}.docx")