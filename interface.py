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
        """

        self.preset_name = preset_name
        preset_data = json.load(open(f'presets/{preset_name}.json'))
    
        self.index = preset_data["index"]
        self.column = preset_data["column"]
        self.value = preset_data["value"]
        self.multisheet = preset_data["multisheet"]
        self.group_by = preset_data["group_by"]
        self.sort = preset_data["sort"]
        self.freeze_panes = (preset_data["freeze_panes"]["row"], preset_data["freeze_panes"]["column"])

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