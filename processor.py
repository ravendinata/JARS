import openpyxl
import pandas as pd

class Processor:
    def __init__(self, source_file_path, output_file_path, mode):
        print("[  ] Initializing processor...")
        
        self.source_file_path = source_file_path
        self.output_file_path = output_file_path
        self.mode = mode
        
        print("[OK] Processor initialized!")

    # Public methods
    def generate_xlsx(self, adjust_cell_widths):
        print(f"[><] Source file path: {self.source_file_path}")
        print(f"[><] Output file path: {self.output_file_path}")
        print(f"[><] Mode: {self.mode}")
    
        print("[..] Processing source file...")
        df = pd.read_csv(self.source_file_path)

        match self.mode:
            case 1:
                print("[  ] Grouping data...")
                grouped = df.groupby("course")
                self.__split_to_sheet(grouped, "Grade", ["student name", "class"], "item name")
            
            case 2:
                print("[  ] Grouping data...")
                grouped = df.groupby("class")
                self.__split_to_sheet(grouped, "Grade", "student name", "course")

            case _:
                raise ValueError("[ER] Invalid report file type.")
            
        # Metadata
        self.add_generic_metadata()

        if adjust_cell_widths:
            self.adjust_cell_widths()

    def add_generic_metadata(self,\
                             title = "",\
                             creator = "JAC Academic Reporting System (JARS)",\
                             subject = "JARS Report",\
                             keywords = "JARS; Academic Report"):
        print("[  ] Adding metadata...")
        workbook = openpyxl.load_workbook(self.output_file_path)
        workbook.properties.title = title
        workbook.properties.creator = creator
        workbook.properties.subject = subject
        workbook.properties.keywords = keywords
        
        workbook.save(self.output_file_path)
        print("[OK] Metadata added!")

    def adjust_cell_widths(self):
        print("[  ] Adjusting cell widths...")
        workbook = openpyxl.load_workbook(self.output_file_path)
        
        for sheet in workbook.worksheets:
            for column_cells in sheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = length + 2
        
        workbook.save(self.output_file_path)
        print("[OK] Cell widths adjusted!")
    
    # Private methods
    def __split_to_sheet(self, grouped_data, value, index, column, sort = False):
        print("[**] Creating output file ...")
        with pd.ExcelWriter(self.output_file_path) as writer:
            for key, item in grouped_data:
                print(f"[  ] Processing {key}...")
                item = grouped_data.get_group(key)
                shaped = item.pivot_table(value, index, column, sort = sort)
                shaped.to_excel(writer, sheet_name = key)
                print(f"[OK] Sheet for {key} created!")

        print("[OK] Processing complete!")