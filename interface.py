import json

class IPresetFile:
    """
    Interface for presets file

    This class is used to access the presets file in a more convenient way.

    Attributes:
        preset_name (str): The name of the preset file.
        preset_json (dict): The preset file in JSON format.
        index (str): The index of the source file.
        column (str): The column of the source file.
        value (str): The value of the source file.
        multisheet (bool): Whether the output file should be multisheet.
        group_by (str): The column to group by.
        sort (str): The column to sort by.
        freeze_panes (tuple): The row and column to freeze panes.
    """

    def __init__(self, preset_name: str):
        """
        Initialize the preset file.

        Args:
            preset_name (str): The name of the preset file.

        TODO:
            Handle preset file using with statement.
        """

        self.preset_name = preset_name
        self.preset_json = json.load(open(f'presets/{preset_name}.json'))
        
        self.index = self.preset_json["index"]
        self.column = self.preset_json["column"]
        self.value = self.preset_json["value"]
        self.multisheet = self.preset_json["multisheet"]
        self.group_by = self.preset_json["group_by"]
        self.sort = self.preset_json["sort"]
        self.freeze_panes = (self.preset_json["freeze_panes"]["row"], self.preset_json["freeze_panes"]["column"])

    def __str__(self):
        """
        Get the name of the preset file.
        
        Returns:
            str: The name of the preset file.
        """
        return self.preset_name
    
    def __repr__(self):
        """
        Returns a string representation of the object.

        Returns:
            str: A string representation of the object.
        """
        return f"IPresetFile({self.preset_name})"