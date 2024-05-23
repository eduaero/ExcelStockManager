from woocommerce import API
import xlwings as xw
import glob
import os
import json 
import pandas as pd
import numpy as np


def define_true(value):
    # Check if query_string_auth is "True" and assign boolean value accordingly
    if value == "True":
        out = True
    else:
        out = False
    return out

def get_latest_xlsx_file(curr_path):
    # Use glob to find all .xlsx files in the directory
    xlsx_files = glob.glob(os.path.join(curr_path, '*.xlsx'))
    # If there are no .xlsx files, return None
    if not xlsx_files:
        return None
    # Get the latest .xlsx file based on modification time
    latest_file = max(xlsx_files, key=os.path.getmtime)   
    return latest_file



def sign_in(config_dict):

    config_keys = ['url', 'consumer_key', 'consumer_secret', 'wp_api', 'version', 'timeout', 'query_string_auth']
    
    for key in config_keys:
        if key not in config_dict:
            raise KeyError(f"Error: Missing key in config dictionary - {key}")

    # If all keys are present, initialize the config
    url = config_dict['url']
    consumer_key = config_dict['consumer_key']
    consumer_secret = config_dict['consumer_secret']
    wp_api = define_true(config_dict['wp_api'])
    version = config_dict['version']
    timeout = config_dict['timeout']
    query_string_auth = define_true(config_dict['query_string_auth'])


    wcapi = API(
        url = url,
        consumer_key = consumer_key,
        consumer_secret = consumer_secret,
        wp_api = wp_api,
        version = version,
        timeout = timeout,
        query_string_auth = query_string_auth)
    # print(" OK ")
    return wcapi

# def all_none(row):
#     return all(x is None or pd.isna(x) for x in row)

def all_none(row):
    return np.all(pd.isnull(row), axis = 1)

def read_config_from_excel(file_path, sheet_name, range, type = "dict"):

    # Check if type is valid
    if type not in ["dict", "df"]:
        raise ValueError("Invalid type in function read_config_from_excel. Expected 'dict' or 'df'")
    
    # Open the workbook
    wb = xw.Book(file_path)
    # Select the 'config' sheet
    sheet = wb.sheets[sheet_name]
    # Read the data starting from cell B4
    data_range = sheet.range(range).value
    # Close the workbook
    wb.close()

    # Convert the data to a dictionary
    if type == "dict":
        config = {row[0]: row[1] for row in data_range}
    else:
        columns = data_range[0]
        data = data_range[1:]
        config = pd.DataFrame(data, columns=columns)
        config.replace('', None, inplace=True)
        config.replace('', np.nan, inplace=True)
        # config.dropna(how='all')
        
        if len(config) > 1:
            bol = all_none(config)
            config = config[~bol]
   
    return config

def write_data_to_excel(df, file_path, sheet_name, start_cell):
    # Open the workbook
    wb = xw.Book(file_path)
    # Select the sheet
    sheet = wb.sheets[sheet_name]
    
    # Clear data starting from the specified cell
    sheet.range(start_cell).expand('table').clear()
    
    # Write the DataFrame to Excel starting from the specified cell
    sheet.range(start_cell).value = df
    
    # Save and close the workbook
    wb.save()
    wb.close()


def save_json(data, out_name = "out"):
    with open(f"{out_name}.json", 'w') as file:
        json.dump(data, file, indent=4)