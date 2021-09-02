from re import S
from ExcelSchemaValidator import ExcelSchemaValidator

path = './Book1.xlsx'
schema = {
    "DNI": {
        "required": True,
        "type": "str",
        "check": {
            "regexp": "^[0-9]{8}$"
        }
    },
    "Celular": {
        "required": True,
        "type": "string",
        "check": {
            "regexp": "^[0-9]{9}$"
        }
    },
    "a'a": {
        "type": "float",
        "check": {
            "min_value": 3.0,
            "max_value": 24.0
        }
    },
    "piña": {

    },
    "aá": {
        "check": {
            "in": ["fsdf", "fdsf", "4324"]
        }
    },
    "date date": {
        "type": "date"
    }
}

excel_schema_validator = ExcelSchemaValidator(path, schema)
print(excel_schema_validator.validate_excel())

# import unittest

