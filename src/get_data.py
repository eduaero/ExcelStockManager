
import json
import pandas as pd
from src.utils.utils import get_latest_xlsx_file, read_config_from_excel, sign_in, define_true, write_data_to_excel
from src.utils.prod_description import initialize_tool, center_text

def retrieve_cat_info(data):

    data_extracted = []
    for item in data:
        category_id = item.get('id', None)
        name = item.get('name', '')
        slug = item.get('slug', '')
        parent_id = item.get('parent', None)
        description = item.get('description', '')
        display = item.get('display', '')
        image_src = item.get('image', {}).get('src', '') if item.get('image') else ''
        menu_order = item.get('menu_order', 0)
        count = item.get('count', 0)

        data_extracted.append((category_id, name, slug, parent_id, description, display, image_src, menu_order, count))

    # Create a DataFrame
    df = pd.DataFrame(data_extracted, columns=['ID', 'Name', 'Slug', 'Parent ID', 'Description', 'Display', 'Image Source', 'Menu Order', 'Count'])
    return df

def retrieve_prod_info(data):

    data_extracted = []
    for item in data:
        product_id = item.get('id', None)
        name = item.get('name', '')
        regular_price = item.get('regular_price', '')
        stock_quantity = item.get('stock_quantity', 'N/A')
        total_sales = item.get('total_sales', 0)
        category_id = item.get('categories', [{'id': None}])[0]['id'] if 'categories' in item else 'N/A'
        category_name = item.get('categories', [{'name': 'N/A'}])[0]['name'] if 'categories' in item else 'N/A'
        permalink = item.get('permalink', '')
        date_created = item.get('date_created', '')
        
        data_extracted.append((product_id, name, regular_price, stock_quantity, total_sales, category_id, category_name, permalink, date_created))

    # Create a DataFrame
    df = pd.DataFrame(data_extracted, columns=['ID', 'Name', 'Regular Price', 'Stock Quantity', 'Total Sales', 'Category ID', 'Category Name', 'Permalink', 'Date Created'])
    return df



class Excel_prods:
    def __init__(self, curr_path):
        initialize_tool()
        self.curr_path = curr_path
        self.excel_path = get_latest_xlsx_file(self.curr_path)
        # Read the information to access to your Woocommerce site
        try:
            self.config = read_config_from_excel(file_path=self.excel_path, sheet_name='config', range='B4:C10')
        except PermissionError:
            raise PermissionError(f"- - - - - ERROR: Please close the Excel file <<{self.excel_path}>> and run the program again - - - - - ")

        # Read the Excel running mode
        try:
            self.exec_mode = read_config_from_excel(file_path=self.excel_path, sheet_name='Control Pannel', range='B6:C11')
        except PermissionError:
            raise PermissionError(f"- - - - - ERROR: Please close the Excel file <<{self.excel_path}>> and run the program again - - - - - ")

        # Initialization of config
        self.initialize_config()

        # Singing-in in Woocommerce
        self.wcapi = sign_in(self.config)
        print("=== CONNECTED: The tool has been connected to Woocommerce")


    def refresh_prod_excel(self):
        self.prod_info = retrieve_prod_info(self.wcapi.get("products").json())
        write_data_to_excel(df = self.prod_info, file_path=self.excel_path, sheet_name='prod_info', start_cell='B5')
        self.cat_info = retrieve_cat_info(self.wcapi.get("products/categories").json())
        write_data_to_excel(df = self.cat_info, file_path=self.excel_path, sheet_name='cat_info', start_cell='B5')
        print("=== RETRIEVING INFO: information retrieved from products and categories")
  

    def initialize_config(self):
        config_keys = ['url', 'consumer_key', 'consumer_secret', 'wp_api', 'version', 'timeout', 'query_string_auth']
        
        for key in config_keys:
            if key not in self.config:
                raise KeyError(f"Error: Missing key in config dictionary - {key}")

        # If all keys are present, initialize the config
        self.url = self.config['url']
        self.consumer_key = self.config['consumer_key']
        self.consumer_secret = self.config['consumer_secret']
        self.wp_api = define_true(self.config['wp_api'])
        self.version = self.config['version']
        self.timeout = self.config['timeout']
        self.query_string_auth = define_true(self.config['query_string_auth'])



