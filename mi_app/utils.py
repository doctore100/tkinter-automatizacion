import os
import json
import datetime
import pandas as pd
# Path utilities
def get_default_template_path():
    """Return the path to the default template file"""
    return os.path.join('templates', 'default_template.docx')

def get_credentials_path():
    """Return the path to the credentials file"""
    return 'credentials.json'

# Validation utilities
def validate_json_file(file_path, required_fields=None):
    """
    Validate that a file exists, is a valid JSON file, and contains required fields.
    
    Args:
        file_path: Path to the JSON file
        required_fields: List of field names that must be present in the JSON
        
    Returns:
        (bool, str): Tuple containing (is_valid, message)
    """
    if not os.path.exists(file_path):
        return False, f"File not found: {file_path}"
    
    try:
        with open(file_path, 'r') as f:
            json_data = json.load(f)
            
        if required_fields:
            for field in required_fields:
                if field not in json_data:
                    return False, f"Invalid JSON: missing {field}"
                    
        return True, "JSON file is valid"
    except json.JSONDecodeError:
        return False, "Invalid JSON format"
    except Exception as e:
        return False, f"Error validating JSON: {str(e)}"


def clean_data(data_fields: dict, dataframes) -> dict:
    """
    Cleans the data by replacing NaN values with empty strings and extracting values from dataframes.

    Args:
        data_fields (dict): A dictionary containing field names as keys and tuples of (row, col) indices as values.

    Returns:
        dict: A cleaned dictionary with field names as keys and corresponding values from dataframes,
              where NaN values are replaced with empty strings.
    """
    cleaned_data = {
        field: '' if pd.isna(dataframes.iloc[row, col]) else dataframes.iloc[row, col]
        for field, (row, col) in data_fields.items()
    }
    return cleaned_data