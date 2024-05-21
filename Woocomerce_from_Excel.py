
from src.get_data import Excel_prods
import os

curr_path = os.path.dirname(os.path.abspath(__file__))
Excel = Excel_prods(curr_path)
Excel.refresh_prod_excel()

# Read the configuration

print("ok")