import json

class IPresetFile:
    def __init__(self, preset_name):
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
        return self.preset_name
    
    def __repr__(self):
        return f"IPresetFile({self.preset_name})"