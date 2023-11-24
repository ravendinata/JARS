import openpyxl
import pandas as pd

import interface

class Formatter:
    """
    Report formatter class for JARS data processor

    This class is responsible for processing the database-like table structured CSV source file and generating the output XLSX file.

    Attributes:
        source_file_path (str): The path of the source file.
        output_file_path (str): The path of the output file.
        preset (interface.IPresetFile): The preset settings information.
        data (pandas.DataFrame): The data from the source file.
    """

    def __init__(self, source_file_path, output_file_path, preset):
        """
        Initialize the processor instance.

        Args:
            source_file_path (str): The path of the source file.
            output_file_path (str): The path of the output file.
            preset (interface.IPresetFile): The name of the preset file to use.

        Returns:
            Processor: The initialized processor instance.
        """
        
        print("[  ] Initializing processor...")
        
        self.source_file_path = source_file_path
        self.output_file_path = output_file_path
        self.preset = interface.IFormatterPresetFile(preset)
        
        print("[OK] Processor initialized!")

    # Public methods
    def generate_xlsx(self, adjust_cell_widths):
        """
        Generate the output XLSX file from the source file data.

        Using the data from the source file and the preset settings, generate the output XLSX file.
        It first reads the source file into a pandas DataFrame, then shapes the data according to the preset settings.
        If the preset settings indicate that the output file should be multisheet, it will split the data into multiple sheets.
        Otherwise, it will shape the data into a single sheet. Then, it will export the data into an XLSX file.
        It will also add a generic metadata to the XLSX file. And optionally adjust the cell widths.

        Args:
            adjust_cell_widths (bool): Whether to adjust cell widths.

        Returns:
            None
        """

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
        """
        Add generic metadata to the output XLSX file.

        Injects generic metadata to the output XLSX file. The metadata includes the title, creator, subject, and keywords.
        If the metadata is not provided, it will use the default values.
        
        Args:
            title (str): The title of the XLSX file.
            creator (str): The creator of the XLSX file.
            subject (str): The subject of the XLSX file.
            keywords (str): The keywords of the XLSX file.
            
        Returns:
            None
        """

        print("[  ] Adding metadata...")
        workbook = openpyxl.load_workbook(self.output_file_path)
        workbook.properties.title = title
        workbook.properties.creator = creator
        workbook.properties.subject = subject
        workbook.properties.keywords = keywords
        
        workbook.save(self.output_file_path)
        print("[OK] Metadata added!")

    def adjust_cell_widths(self):
        """
        Adjust the cell widths of the output XLSX file.

        It will adjust the cell widths of all sheets in the XLSX file according to the content of the cells.

        Returns:
            None
        """

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
        """
        Split the data into multiple sheets.

        Using the preset settings, it will split the data into multiple sheets within the same XLSX file.
        The split data is also shaped according to the preset settings. Then, it will export the data into an XLSX file.

        Returns:
            None
        """

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
        """
        Shape the data and export it into an XLSX file.
        
        Using the preset settings, it will shape the data into a single sheet within a single XLSX file and export it.
        
        Returns:
            None
        """

        print(f"[**] Creating output file...")
        shaped = self.data.pivot_table(self.preset.value, self.preset.index, self.preset.column, sort = self.preset.sort)
        shaped.to_excel(self.output_file_path, sheet_name = self.preset.preset_name, freeze_panes = self.preset.freeze_panes)
        print(f"[OK] Sheet for {self.preset} created!")