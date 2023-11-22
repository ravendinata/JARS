import interface
import openpyxl
import pandas as pd

class Processor:
    def __init__(self, source_file_path, output_file_path, preset):
        print("[  ] Initializing processor...")
        
        self.source_file_path = source_file_path
        self.output_file_path = output_file_path
        self.preset = interface.IPresetFile(preset)
        
        print("[OK] Processor initialized!")

    # Public methods
    def generate_xlsx(self, adjust_cell_widths):
        print(f"[><] Source file path: {self.source_file_path}")
        print(f"[><] Output file path: {self.output_file_path}")
        print(f"[><] Preset: {self.preset}")
    
        print("[..] Processing source file...")
        df = pd.read_csv(self.source_file_path)

        if self.preset.multisheet:
            print("[  ] Grouping data...")
            self.data = df.groupby(self.preset.group_by)
            self.__split_to_sheet()
        else:
            self.data = df
            self.__shape_and_export()

        # Metadata
        self.add_generic_metadata()

        if adjust_cell_widths:
            self.adjust_cell_widths()

    def add_generic_metadata(self,
                             title = "",
                             creator = "JAC Academic Reporting System (JARS)",
                             subject = "JARS Report",
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
    def __split_to_sheet(self):
        print("[**] Creating output file...")

        with pd.ExcelWriter(self.output_file_path) as writer:
            for key, item in self.data:
                print(f"[  ] Processing {key}...")
                item = self.data.get_group(key)
                shaped = item.pivot_table(self.preset.value, self.preset.index, self.preset.column, sort = self.preset.sort)
                shaped.to_excel(writer, sheet_name = key, freeze_panes = self.preset.freeze_panes)
                print(f"[OK] Sheet for {key} created!")

        print("[OK] Processing complete!")

    def __shape_and_export(self):
        print(f"[**] Creating output file...")
        shaped = self.data.pivot_table(self.preset.value, self.preset.index, self.preset.column, sort = self.preset.sort)
        shaped.to_excel(self.output_file_path, sheet_name = self.preset.preset_name, freeze_panes = self.preset.freeze_panes)
        print(f"[OK] Sheet for {self.preset} created!")