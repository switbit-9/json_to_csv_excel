import json
import os
import sys
from parse import CSVConverter, ExcelConverter
from dotenv import load_dotenv
load_dotenv()

TARGET_FORMAT = ''

def check_target():
    global TARGET_FORMAT
    TARGET_FORMAT = os.getenv('TARGET_FORMAT')
    if TARGET_FORMAT == 'csv':
        print("Converting file to csv ...")
    elif TARGET_FORMAT == 'excel':
        print('Converting file to excel ...')

    else:
        print(f'This {TARGET_FORMAT} format is not supported !!')
        sys.exit()

def load_input_file():
    input_file = os.getenv('INPUT_FILE')
    input_file_path = os.path.join(os.getcwd(), input_file)
    try:
        with open(input_file_path, 'r') as file:
            data = json.load(file)
            return data
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        sys.exit()
    except FileNotFoundError as e:
        print(f"File not found on specified path: {input_file}")
        sys.exit()


def main():
    check_target()
    data = load_input_file()
    for param in data:
        filename = param['FileName']

        if param['FamilyLoaded'] is False:
            print('This Filename is skipped because is not loaded ...')
            continue
        params_info = param['ParameterInformation']

        if TARGET_FORMAT == 'csv':
            for item in params_info:
                parser = CSVConverter(item)
            parser.write_to_file(filename)

        elif TARGET_FORMAT == 'excel':
            for item in params_info:
                parser = ExcelConverter(item)
            parser.write_to_file(filename)














if __name__ == '__main__':
    main()