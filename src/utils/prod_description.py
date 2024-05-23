# Function to center-align text
def center_text(text, width = 70):
    return text.center(width)


# Welcome message for ExcelStockManager
def initialize_tool():
    print("=" * 70)
    print(center_text("Welcome to: ExcelStockManager"))
    print(center_text("A Python tool for managing product stock in WooCommerce from Excel"))
    print(center_text("Developed by: Eduardo Tola Losa"))
    print(center_text("LinkedIn: https://www.linkedin.com/in/eduardotolalosa/"))
    print(center_text("GitHub: https://github.com/eduaero/ExcelStockManager"))
    print("=" * 70)