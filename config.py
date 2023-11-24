import json

@staticmethod
def get_config(config_name):
    with open('config.json') as json_file:
        data = json.load(json_file)
        return data[config_name]