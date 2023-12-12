from datetime import datetime

from pypdf import PdfWriter, PdfReader

@staticmethod
def pdf_inject(file_path, student_name, grader_report, time_docsaved = datetime.now()):
    # Add metadata to PDF
    reader = PdfReader(file_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.add_metadata(
        {
            "/Author": "JAC Academic Reporting System",
            "/Title": f"{student_name} - {grader_report.get_course_info('Subject')} - S{grader_report.get_course_info('Semester')} AY{grader_report.get_course_info('School Year')} Report Card",
            "/Category": "Semester Report Card",
            "/Revision": 1,
            "/Keywords": "JAC; JARS; Report Card; Semester Report Card",
            "/Generated Time": time_docsaved.strftime(f"%Y/%m/%d @ %H:%M:%S"),
            "/CreationDate": datetime.now().strftime(f"D\072%Y%m%d%H%M%S%z"),
            "/Creator": "JARS/CReP 1.0.0",
            "/Producer": "pypdf2 @ JARS/CReP 1.0.0",
        }
    )

    with open(file_path, "wb") as output:
        writer.write(output)