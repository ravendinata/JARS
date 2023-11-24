import json

class IFormatterPresetFile:
    """
    Interface for presets file

    This class is used to access the presets file in a more convenient way.

    Attributes:
        preset_name (str): The name of the preset file.
        index (int): The index of the column to use as the index.
        column (int): The index of the column to use as the column.
        value (int): The index of the column to use as the value.
        multisheet (bool): Whether to generate a multisheet file.
        group_by (int): The index of the column to use as the group by column.
        sort (bool): Whether to sort the data.
        freeze_panes (tuple): The coordinates of the cell to freeze panes at.

    Methods:
        __init__(self, preset_name): Initialize the preset file.
        __str__(self): Get the name of the preset file.
        __repr__(self): Returns a string representation of the object.
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