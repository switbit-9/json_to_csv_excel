import os
import csv
import sys
import pandas as pd
import re
class BaseConverter:
    FIRST_ITER = True
    data = {
    }

    column_names = {
        "Merkmalsname" : "Merkmalsname",
        "datetyp_bms": "Datentyp BMS",
        "revit_internal_name" : "Revit Internal Name",
        "parameter_type" : "Parameter Type (UnitType)",
        "display_unit_type" : "DisplayUnitType",
        "family_shared_parameter" : "Family-/Shared-Parameter",
        "type_instance_parameter" : "Type-/Instance-Parameter",
        "parameter_section" : "Parameter-Section (Revit)",
        "GuidDe" : "GuidDe"
    }

    def __init__(self, item):
        self.filename = ''
        self.Merkmalsname = item.get('ParameterName', 'No Value')
        parameterValue = item.get('ParameterValues', None)
        if parameterValue is not None:
            self.datetyp_bms = self.check_parameter_values_type(parameterValue['Allgemein']) if parameterValue.get('Allgemein', None) is not None else "No Value"
        else:
            self.datetyp_bms = 'No Value'
        self.revit_internal_name = item.get('BuiltInParameterEnumName', 'No Value')
        self.parameter_type = item.get('ParameterValueType', 'No Value') #unittype
        self.display_unit_type = self.check_display_unit_type(item)

        self.family_shared_parameter = self.check_family_shared_param(item)
        self.type_instance_parameter = self.check_type_instance_param(item)
        self.parameter_section = item.get('ParameterGroup', 'No Value')
        self.GuidDe = item.get('Guid', 'No Value')

        self.convert()

    def check_parameter_values_type(self, value):
        if value in ['0', '1', True, False]:
            return "Bool"
        elif re.match(r'^[+-]?\d+(\.\d+)?$', value):
            return 'Decimal' if '.' in value else 'Integer'
        elif isinstance(value, str):
            return "String"
        else:
            return 'Not Specified'
    def check_display_unit_type(self, item):
        unit_values = ['BSU_CUBIC_METERS_PER_HOUR', 'BSU_MILLIMETERS', "BSU_CURRENCY", "BSU_WATTS", "BSU_VOLTS", 'BSU_VOLT_AMPERES', 'BSU_GENERAL', "BSU_PASCALS"]
        unit = item.get('Unit', 'NO_UNIT')
        return unit if unit in unit_values else "NO_UNIT"

    def check_family_shared_param(self, item):
        param = item.get('IsSharedParameter', 'No Value')
        if isinstance(param, bool):
            param = 'Shared' if param is True else 'Family'
        else:
            param = 'Wrong Value Added !'
        return param

    def check_type_instance_param(self, item):
        param = item.get('IsInstanceParameter', 'No Value')
        if isinstance(param, bool):
            param = 'Instance' if param is True else 'Type'
        else:
            print("Wrong Value Added !")
        return param

    def check_target_folder(self, filename):
        self.filename = filename.split('.')[0] + self.file_extention
        directory_path = os.path.join(os.getcwd(), self.target_directory)
        if not os.path.exists(self.target_directory) and not os.path.isdir(self.target_directory):
            os.makedirs(directory_path)
        return os.path.join(directory_path, self.filename)

    def convert(self):
        properties = self.column_names
        if self.FIRST_ITER is True:
            for item in properties.keys():
                self.data[item] = []
            BaseConverter.FIRST_ITER = False


        for item in properties.keys():
            self.data[item].append(self.__dict__[item])




class ExcelConverter(BaseConverter):
    file_extention = '.xlsx'
    target_directory = 'target_excel'

    def write_to_file(self, filename):
        file_path = self.check_target_folder(filename)
        try:
            df_existing = pd.read_excel(file_path)
            df_new = pd.DataFrame(self.data)
            df_combined = pd.concat([df_existing, df_new])
            df_combined.to_excel(file_path)

        except (pd.errors.EmptyDataError, FileNotFoundError):
            df = pd.DataFrame(self.data)
            df.rename(columns=self.column_names, inplace=True)
            df.to_excel(file_path, index=False)

        except Exception as e:
            print(e)
            sys.exit()



class CSVConverter(BaseConverter):
    file_extention = '.csv'
    target_directory = 'target_csv'

    def write_to_file(self, filename):
        file_path = self.check_target_folder(filename)

        try:
            old_df = pd.read_csv(file_path)
            new_df = pd.DataFrame(self.data)
            combined_df = pd.concat([old_df, new_df]).to_csv(filename, index=False)


        except (pd.errors.EmptyDataError, FileNotFoundError):
            df = pd.DataFrame(self.data)
            df.rename(columns=self.column_names, inplace=True)
            df.to_csv(file_path, index=False)

        except Exception as e:
            print(e)
            sys.exit()