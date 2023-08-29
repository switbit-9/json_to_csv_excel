import os
import csv
import sys
import pandas as pd
import re

class BaseConverter():
    file_format = 'csv'
    FIRST_ITER = True
    data = {}
    file_path = ''

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
        self.Merkmalsname = item.get('ParameterName', 'No Value')


        self.datetyp_bms = item['ParameterValues'].get('Allgemein', False) if item.get('ParameterValues', False) else False


        self.revit_internal_name = item.get('BuiltInParameterEnumName', 'No Value')
        self.parameter_type = item.get('ParameterValueType', 'No Value') #unittype
        # self.display_unit_type = self.check_display_unit_type(item)

        self.display_unit_type = item.get('Unit', False)
        # self.family_shared_parameter = self.check_family_shared_param(item)
        self.family_shared_parameter = item.get('IsSharedParameter', False)
        # self.type_instance_parameter = self.check_type_instance_param(item)
        self.type_instance_parameter = item.get('IsInstanceParameter', False)
        self.parameter_section = item.get('ParameterGroup', 'No Value')
        self.GuidDe = item.get('Guid', 'No Value')

        self.convert()

    @property
    def type_instance_parameter(self):
        return self._type_instance_parameter

    @type_instance_parameter.setter
    def type_instance_parameter(self, value):
        if isinstance(value, bool):
           param = 'Instance' if value is True else 'Type'
        elif not value:
            param = 'No Value'
        else:
            print("Wrong Value Added !")
            param = 'Wrong Value !'
        self._type_instance_parameter = param


    @property
    def family_shared_parameter(self):
        return self._family_shared_parameter

    @family_shared_parameter.setter
    def family_shared_parameter(self, value):
        if isinstance(value, bool):
            param = 'Shared' if value is True else 'Family'
        elif not value:
            param = 'No Value'
        else:
            print("Wrong Value added !")
            param = 'Wrong Value!'
        self._family_shared_parameter = param


    @property
    def display_unit_type(self):
        return self._display_unit_type

    @display_unit_type.setter
    def display_unit_type(self, value):
        if not value:
            return 'No Value'
        unit_values = ['BSU_CUBIC_METERS_PER_HOUR', 'BSU_MILLIMETERS', "BSU_CURRENCY", "BSU_WATTS", "BSU_VOLTS", 'BSU_VOLT_AMPERES', 'BSU_GENERAL', "BSU_PASCALS"]
        self._display_unit_type = value if value in unit_values else "NO_UNIT"

    @property
    def datetyp_bms(self):
        return self._datetyp_bms

    @datetyp_bms.setter
    def datetyp_bms(self, value):
        if not value:
            param = 'No Value'
        elif value in ['0', '1', True, False]:
            param = "Bool"
        elif re.match(r'^[+-]?\d+(\.\d+)?$', value):
            param = 'Decimal' if '.' in value else 'Integer'
        elif isinstance(value, str):
            param = "String"
        else:
            param = 'No Value'
        self._datetyp_bms = param

    def check_target_file(self, filename, to_format, target_directory):
        self.set_filename(to_format, filename)
        target_directory = self.set_target_directory(target_directory)
        directory_path = os.path.join(os.getcwd(), target_directory)
        if not os.path.exists(target_directory) and not os.path.isdir(target_directory):
            os.makedirs(directory_path)
        self.set_file_path(os.path.join(directory_path, self.filename))

    @classmethod
    def set_file_path(cls, file_path):
        cls.file_path = file_path
    @classmethod
    def modify_first_iter(cls, value=False):
        cls.FIRST_ITER = value

    @classmethod
    def quit(cls):
        cls.FIRST_ITER = True
        cls.data = {}



    def convert(self):
        properties = self.column_names
        if self.FIRST_ITER is True:
            for item in properties.keys():
                self.data[item] = []
            self.modify_first_iter()

        for item in self.data.keys():
            if item in self.__dict__.keys():
                self.data[item].append(self.__dict__[item])
            else:
                self.data[item].append(self.__dict__["_" + item])

    @classmethod
    def set_filename(cls, to_format, filename):
        if to_format == 'csv':
            cls.file_format = 'csv'
        elif to_format == 'excel':
            cls.file_format = 'xlsx'
        cls.filename = filename.split('.')[0] + '.' + cls.file_format

    def set_target_directory(self, target_directory=None):
        if target_directory is None:
            if self.file_format == 'csv':
                return 'target_csv'
            elif self.file_format == 'xlsx':
                return 'target_excel'
        return target_directory

class Converter(BaseConverter):
    def __init__(self, items, filename, to_format='csv', target_directory=None):
        if self.FIRST_ITER == True:
            self.check_target_file(filename, to_format, target_directory)
        for item in items:
            super().__init__(item)
        self.write_to_file()


    def write_to_file(self):
        # file_path = self.check_target_folder(filename)
        file_path = self.file_path
        try:
            old_df = pd.read_csv(file_path) if self.file_format ==  'csv' else pd.read_excel(file_path)
            new_df = pd.DataFrame(self.data)
            df = pd.concat([old_df, new_df])
            df.to_csv(file_path, index=False) if self.file_format == 'csv' else df.to_excel(file_path, index=False)
        except (pd.errors.EmptyDataError, FileNotFoundError):
            df = pd.DataFrame(self.data)
            df.rename(columns=self.column_names, inplace=True)
            df.to_csv(file_path, index=False) if self.file_format == 'csv' else df.to_excel(file_path, index=False)
        except Exception as e:
            print(e)
            sys.exit()
        finally:
            self.quit()



