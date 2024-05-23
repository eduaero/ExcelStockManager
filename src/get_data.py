
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

def process_cat_df(base_url, data):
    # Filter categories with Upload = 'Yes'
    filtered_data = data[data['Upload'] == 'Yes'].copy()
    
    # Delete unnecessary columns
    filtered_data.drop(columns=['Activate', 'Upload'], inplace=True)
    
    # Handle 'image' column
    for index, row in filtered_data.iterrows():
        if pd.notnull(row['image']):  # Check if 'image' is not None
            img_path = row['image']
            # Check if the image path is complete or relative
            if not (img_path.startswith('http://') or img_path.startswith('https://')):
                img_path = f"{base_url}/wp-content/uploads/{img_path}"
            filtered_data.at[index, 'image'] = {'src': img_path}

    categories_data = [{k: v for k, v in row.items() if v is not None} for i, row in filtered_data.iterrows()]
    # print(categories_data)
    return categories_data

def process_product_df(base_url, data):
    # Filter products with Upload = 'Yes'
    filtered_data = data[data['Upload'] == 'Yes'].copy()
    
    # Delete unnecessary columns
    filtered_data.drop(columns=['Activate', 'Upload'], inplace=True)
    
    # Handle images - consider multiple image columns (images1, images2, ...)
    for index, row in filtered_data.iterrows():
        images = []
        for col in data.columns:
            if col.startswith('images') and pd.notnull(row[col]):
                img_path = row[col]
                # Check if the image path is complete or relative
                if not (img_path.startswith('http://') or img_path.startswith('https://')):
                    img_path = f"{base_url}/wp-content/uploads/{img_path}"
                images.append({'src': img_path})
        filtered_data.at[index, 'images'] = images if images else None

    # Handle attributes - consider multiple attribute columns (attributes1, attributes2, ...)
    for index, row in filtered_data.iterrows():
        attributes = []
        for col in data.columns:
            if col.startswith('attributes') and pd.notnull(row[col]):
                attribute_data = row[col].split(':')
                if len(attribute_data) == 2:
                    attr_name, attr_options = attribute_data
                    attributes.append({
                        'name': attr_name.strip(),
                        'options': [opt.strip() for opt in attr_options.split(',')]
                    })
        filtered_data.at[index, 'attributes'] = attributes if attributes else None

    # Creating JSON data
    products_data = []
    for _, row in filtered_data.iterrows():
        product = {k: v for k, v in row.items() if v is not None}
        products_data.append(product)
    
    return products_data


def post_new_cat(wcapi, data):
    failed_count = 0
    correct = 0
    total = len(data)
    for d in data:
        response = (wcapi.post("products/categories", d).json())
        category_name = data["name"]
    
        if "id" not in response.keys():
            failed_count = failed_count + 1
        else:
            correct = correct + 1
            print(f"====== Category {category_name} created successfully.")

    if failed_count == 0:
        print(f"=== CATEGORIES UPLOADED: {correct} categories uploaded correcly")
    else:
        print(f"=== ERROR UPLOADING CATEGORIES: Fail to upload {failed_count} categories out of {total}")

def post_new_product(wcapi, data):
    failed_count = 0
    correct = 0
    total = len(data)
    
    for product in data:
        response = wcapi.post("products", product).json()
        category_name= data["name"]
        
        if "id" not in response.keys():
            failed_count += 1
            # print(wcapi.post("products", product).json())
        else:
            correct = correct +1
            print(f"====== Product {category_name} created successfully.")

    
    if failed_count == 0:
        print(f"=== PRODUCTS UPLOADED: {correct} Product uploaded correctly")
    else:
        print(f"=== ERROR UPLOADING PRODUCTS: Fail to upload {failed_count} products out of {total}")

def update_product_stock(wcapi, filtered_data):
    for index, row in filtered_data.iterrows():
        product_id = int(row['ID'])
        if pd.notnull(row["New stock"]):
            data = {'stock_quantity': int(row["New stock"])}
            response = wcapi.put(f"products/{product_id}", data).json()
            if 'id' in response:
                print(f"====== Product {product_id} stock updated successfully to {row["New stock"]}")
            else:
                print(f"Failed to update product {product_id} stock. Response: {response}")

def update_product_prices(wcapi, filtered_data):
    for index, row in filtered_data.iterrows():
        product_id = int(row['ID'])
        if pd.notnull(row["New price"]):
            data = {'regular_price': str(row["New price"])}
            response = wcapi.put(f"products/{product_id}", data).json()
            if 'id' in response:
                print(f"====== Product {product_id} price updated successfully to {row["New price"]}")
            else:
                print(f"Failed to update product {product_id} price. Response: {response}")


class Excel_prods:
    def __init__(self, curr_path):
        initialize_tool()
        self.curr_path = curr_path
        self.excel_path = get_latest_xlsx_file(self.curr_path)

        self.read_excel_info()

        # Initialization of config
        self.initialize_config()

        # Singing-in in Woocommerce
        self.wcapi = sign_in(self.config)
        print("=== CONNECTED: The tool has been connected to Woocommerce")

    def check_config(self, detail_df, mode):
        if mode == 0:
            # Mode 0: No need to check
            return True
        elif mode == 1:
            # Mode 1: There should be at least 1 new category
            new_categories = detail_df[(detail_df['Description'] == 'Categories to upload') & (detail_df['To upload'] > 0)]
            if not new_categories.empty:
                return True
        elif mode == 2:
            # Mode 2: There should be at least 1 new product and no new categories
            new_products = detail_df[(detail_df['Description'] == 'Products to upload') & (detail_df['To upload'] > 0)]
            new_categories = detail_df[(detail_df['Description'] == 'Categories to upload') & (detail_df['To upload'] > 0)]
            if not new_products.empty and new_categories.empty:
                return True
        elif mode == 3:
            # Mode 3: No new categories and products, at least 1 price or stock to update
            new_categories = detail_df[(detail_df['Description'] == 'Categories to upload') & (detail_df['To upload'] > 0)]
            new_products = detail_df[(detail_df['Description'] == 'Products to upload') & (detail_df['To upload'] > 0)]
            new_prices_or_stock = detail_df[(detail_df['Description'].isin(['Prices to update', 'Stock numbers to update'])) | (detail_df['To upload'] > 0)]
            if new_categories.empty and new_products.empty and not new_prices_or_stock.empty:
                return True
        return False

    def desc_exec_mode(self):
        # Converting the DataFrame to a dictionary
        df_dict = self.exec_mode.to_dict(orient='records')

        
        # Check each execution mode

        self.mode = df_dict[0]['Mode']
        self.mode_name = df_dict[0]['Desc']
        is_configured_correctly = self.check_config(self.exec_sum, self.mode)
        
        print(f"=== EXECUTION MODE: {self.mode} - Executing to {self.mode_name}. {'Configured correctly' if is_configured_correctly else f'Configuration error. The tool cannot be configures in Mode: {self.mode} ==='}")

    def Excel_execute(self):

        self.desc_exec_mode()

        # 0. Always refresh Excel information at the begginig of the execution
        self.refresh_prod_excel()

        if self.mode == 1: # Execute to upload categories
            data = process_cat_df(self.config["url"], self.new_cat)
            post_new_cat(self.wcapi, data)
        elif self.mode == 2:
            data = process_product_df(self.config["url"], self.new_prod)
            post_new_product(self.wcapi, data)
        elif self.mode == 3 or self.mode == 4:

            # Decide whether to update stock or price based on mode
            data_df = self.prod_price_stock[self.prod_price_stock['Activate'] == 'Yes']
            if self.mode == 3:
                update_product_stock(self.wcapi, data_df)
            elif self.mode == 4:
                update_product_prices(self.wcapi, data_df)
        else:
            print(f"=== ERROR: Unsupported mode: {self.mode}")

        # 99. Always refresh Excel information at the end of the execution
        self.refresh_prod_excel()

        print("=== EXECUTION FINISHED!")
        print("=" * 70)

    def read_excel_info(self):

        error_msg = f"- - - - - ERROR: Please close the Excel file <<{self.excel_path}>> and run the program again - - - - - "

        try:
            # 0. Get all the location of the input data:
            self.info_source = read_config_from_excel(file_path = self.excel_path, sheet_name='99. Information source', range='B4:K33', type = "df")

            # Initialize a list to store data for each row as a dictionary
            table_data = []

            # Iterate over each row in the DataFrame
            for index, row in self.info_source.iterrows():
                row_data = {}
                for column in self.info_source.columns:
                    row_data[column] = row[column]
                table_data.append(row_data)

            for t in table_data:

                data = read_config_from_excel(file_path = self.excel_path, sheet_name = t["sheet"], range = t["array"], type = t["type"])
                setattr(self, t["py_name"], data)

        except PermissionError:
            raise PermissionError(error_msg)
        

    def refresh_prod_excel(self):
        products = []
        page = 1
        while True:
            response = self.wcapi.get("products", params={"per_page": 100, "page": page})
            if response.status_code != 200:
                print(f"Error fetching products: {response.status_code} - {response.text}")
                break
            data = response.json()
            if not data:
                break
            products.extend(data)
            page += 1
        self.prod_info = retrieve_prod_info(products)
        write_data_to_excel(df = self.prod_info, file_path=self.excel_path, sheet_name='prod_info', start_cell='B5')

        # Retrieve info from the categories
        categories = []
        page = 1
        while True:
            response = self.wcapi.get("products/categories", params={"per_page": 100, "page": page})
            if response.status_code != 200:
                print(f"Error fetching categories: {response.status_code} - {response.text}")
                break
            data = response.json()
            if not data:
                break
            categories.extend(data)
            page += 1

        self.cat_info =  retrieve_cat_info(categories)
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



