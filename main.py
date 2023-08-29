import json
import os
import sys
from parse import Converter
from dotenv import load_dotenv
load_dotenv()


def parse_args():
    target_format = os.getenv('TARGET_FORMAT')
    if target_format == 'csv':
        print("Converting file to csv ...")
    elif target_format == 'excel':
        print('Converting file to excel ...')
    else:
        print(f'This {target_format} format is not supported !!')
        sys.exit()

    input_file_path = os.path.join(os.getcwd(), os.getenv('INPUT_FILE'))
    return target_format, input_file_path

def load_input_file(file_path):
    try:
        with open(file_path, 'r') as file:
            data = json.load(file)
            return data
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        sys.exit()
    except FileNotFoundError as e:
        print(f"File not found on specified path: {file_path}")
        sys.exit()


def main():
    target_format, file_path = parse_args()
    data = load_input_file(file_path)
    # converter = CSVConverter if TARGET_FORMAT == 'csv' else ExcelConverter
    for param in data:
        filename = param['FileName']
        if param['FamilyLoaded'] is False:
            print(f'This {filename} is skipped because is not loaded ...')
            continue
        params_info = param['ParameterInformation']
        Converter(params_info, filename, to_format=target_format)


if __name__ == '__main__':
    main()