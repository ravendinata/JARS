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
    def generate_xlsx(self):
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
        print("[  ] Adding metadata...")
        workbook = openpyxl.load_workbook(self.output_file_path)
        workbook.properties.creator = "JAC Academic Reporting System (JARS)"
        workbook.properties.subject = "JARS Report"
        workbook.properties.keywords = "JARS; Academic Report"
        
        workbook.save(self.output_file_path)
        print("[OK] Metadata added!")
    
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