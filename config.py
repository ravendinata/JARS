import json

@staticmethod
def get_config(config_name):
    with open('config.json') as json_file:
        data = json.load(json_file)
        return data[config_name]
    
@staticmethod
def set_config(config_name, config_value):
    with open('config.json') as json_file:
        data = json.load(json_file)
        data[config_name] = config_value
        with open('config.json', 'w') as json_file:
            json.dump(data, json_file, indent = 4)