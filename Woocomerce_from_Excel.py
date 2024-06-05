
from src.get_data import Excel_prods
import os

curr_path = os.path.dirname(os.path.abspath(__file__))
Excel = Excel_prods(curr_path)
Excel.Excel_execute()

# get_all = self.wcapi.get("").json()
# get_all["routes"]["/wc/v3/products"]["endpoints"][0]["args"]["tax_class"]["enum"]
