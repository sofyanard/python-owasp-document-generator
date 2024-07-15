import json
import os
def append_to_json(data, filename):
    try:
        with open(filename, 'a') as json_file:
            json.dump(data, json_file, indent=4)
            json_file.write('\n')  # Ensure each entry is on a new line
        print(f'Data successfully appended to {filename}')
    except Exception as err:
        print(f'Error appending data to JSON file: {err}')
        
def write_json_file(data, filename):
    try:
        with open(filename, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        print(f'Data successfully write to {filename}')
    except Exception as err:
        print(f'Error writing data to JSON file: {err}')
        
def initialize_json(filename):
    try:
        with open(filename, 'a') as json_file:
            json_file.write('\n')  # Ensure each entry is on a new line
    except Exception as err:
        print(f'Error Initialize JSON file: {err}')

def validate_json(filename):
    data=None
    if os.path.exists(filename):
        # Read the existing data from the file
        with open(filename, 'r') as file:
            try:
                data = json.load(file)
                if not isinstance(data, list):
                    data = []
            except json.JSONDecodeError:
                data = []
    else:
        data = []
    with open(filename, 'w') as file:
        json.dump(data, file, indent=4)
        
def read_json_file(filename: str) -> dict:
    try:
        with open(filename, 'r', encoding="utf-8") as file:
            content = json.load(file)
        return content
    except FileNotFoundError:
        print("File not Found in "+filename)
        return {}
    except Exception as err:
        print(f'Error Reading JSON file: {repr(err)}')
        return {}