import pandas as pd
import re
from datetime import datetime

class ExcelSchemaValidator():

    def __init__(self, path, schema):
        self.df = pd.read_excel(path, dtype=str)
        self.schema = schema

    def is_required(self, scheme_column, index_row, column_name, cell_value):
        required = scheme_column.get("required", False)
        if (required and (cell_value == "" or cell_value is None)):
            msg = "Columna {} es requerida, No hay valor en la fila {}.".format(column_name, index_row + 1)
            raise Exception(msg)
    
    def check_in(self, schema_check, value):
        in_list = schema_check.get("in", None)
        if not in_list is None and not value in in_list:
            return False
        return True

    def check_min_value(self, schema_check, value):
        min_value = schema_check.get("min_value", None)
        if not min_value is None and min_value > value:
            return False
        return True

    def check_max_value(self, schema_check, value):
        max_value = schema_check.get("max_value", None)
        if not max_value is None and max_value < value:
            return False
        return True

    def check_regex(self, schema_check, value):
        regexp = schema_check.get("regexp", None)
        if not regexp is None and re.compile(regexp).search(value) is None:
            return False
        return True

    def is_check(self, scheme_column, index_row, column_name, value):
        data_type = scheme_column.get("type", "string")
        schema_check = scheme_column.get("check", {})
        if data_type in ("int", "float") and not self.check_min_value(schema_check, value):
            min_value = schema_check.get("min_value", None)
            msg = "Columna {} tiene que tener un valor minimo de {}, valor inválido en la fila {}.".format(column_name, min_value, index_row + 1)
            raise Exception(msg)
        elif data_type in ("int", "float") and not self.check_max_value(schema_check, value):
            max_value = schema_check.get("max_value", None)
            msg = "Columna {} tiene que tener un valor maximo de {}, valor inválido en la fila {}.".format(column_name, max_value, index_row + 1)
            raise Exception(msg)
        elif data_type in ("string") and not self.check_regex(schema_check, value):
            regexp = schema_check.get("regexp", None)
            msg = "Columna {} no tiene el formato {}, valor inválido en la fila {}.".format(column_name, regexp, index_row + 1)
            raise Exception(msg)
        elif data_type in ("string") and not self.check_in(schema_check, value):
            in_list = schema_check.get("in", None)
            msg = "Columna {} tiene que contener alguno de los siguientes valores : {}, valor inválido en la fila {}.".format(column_name, in_list, index_row + 1)
            raise Exception(msg)
    
    def convert_type(self, scheme_column, index_row, column_name, cell_value):
        data_type = scheme_column.get("type", "string")
        if data_type == "int":
            match = re.compile("^[0-9]+$").search(cell_value)
            if not match:
                msg = "Columna {} tiene que ser un número entero, valor inválido en la fila {}.".format(column_name, index_row + 1)
                raise Exception(msg)
            return int(match.group())
        elif data_type == "float":
            match = re.compile("^\d+\.\d+$").search(cell_value) or re.compile("^[0-9]+$").search(cell_value)
            if not match:
                msg = "Columna {} tiene que ser un número decimal, valor inválido en la fila {}.".format(column_name, index_row + 1)
                raise Exception(msg)
            return float(match.group())
        elif data_type == "date":
            try:
                return datetime.strptime(cell_value, "%Y-%m-%d %H:%M:%S")
            except Exception as e:
                try:
                    return datetime.strptime(cell_value,"%Y-%m-%d")
                except Exception as e:
                    try:
                        return datetime.strptime(cell_value, "%Y/%m/%d")
                    except Exception as e:
                        try:
                            return datetime.strptime(cell_value, "%Y/%m/%d %H:%M:%S")
                        except Exception as e:
                            msg = "Columna {} tiene que tener formato de fecha, valor inválido en la fila {}.".format(column_name, index_row + 1)
                            raise Exception(msg)
        return cell_value
    
    def validate_excel(self):
        rows = self.df.shape[0]
        table_dict = []
        errors = []
        for index_row in range(0, rows):
            row_dict = {}
            for column_name in self.schema.keys():
                try:
                    scheme_column = self.schema[column_name]
                    required = scheme_column.get("required", False)
                    if column_name in self.df.columns:
                        cell_value = self.df.iloc[index_row][column_name]
                        self.is_required(scheme_column, index_row, column_name, cell_value)
                        value = self.convert_type(scheme_column, index_row, column_name, cell_value)
                        self.is_check(scheme_column, index_row, column_name, value)
                        row_dict[column_name] = value
                    elif required:
                        msg = "Columna {} no encontrada para la fila {}.".format(column_name, index_row + 1)
                        raise Exception(msg)
                    elif not required:
                        row_dict[column_name] = None
                except Exception as e:
                    errors.append(str(e))
            table_dict.append(row_dict)
        return table_dict, errors