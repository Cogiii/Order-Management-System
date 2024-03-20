import openpyxl 
import os

class Resources:
    def __init__(self):
        # COLOR PALLETE FROM THIS LINK
        # https://colorhunt.co/palette/5356ff378ce767c6e3dff5ff
        self.COLOR1 = "#5356FF"
        self.COLOR2 = "#378CE7"
        self.COLOR3 = "#67C6E3"
        self.COLOR4 = "#DFF5FF"

        self.currency_sign = "₱"
        self.table_header = ('Order ID','Table Number','Orders','Date','Bill','Status')
        self.workbook = None
        self.sheet = None
        self.menu_price = None

    def open_file(self):
        self.EXCEL_FILE = 'database/Orders.xlsx'

        # Check if the Excel file exists
        if os.path.exists(self.EXCEL_FILE):
            try:
                # Load workbook and get the data only (ignore formulas)
                self.workbook = openpyxl.load_workbook(self.EXCEL_FILE, data_only=True)
                self.sheet = self.workbook.active

            except Exception as e:
                print(f"An error occurred while loading the workbook: {e}")
        else:
            print(f"The Excel file '{self.EXCEL_FILE}' does not exist.")

        # Get Price from excel
        self.menu_price = {"Pizza": self.sheet["J2"].value, 
                            "Carbonara": self.sheet["J3"].value,
                            "Chicken": self.sheet["J4"].value,
                            "Pancit": self.sheet["J5"].value,
                            "Lasagna": self.sheet["J6"].value}
    
    def center_window(self, Screen, width, height, scale_factor=1.0):
        screen_width = Screen.winfo_screenwidth()
        screen_height = Screen.winfo_screenheight()
        
        x = int(((screen_width/2) - (width/2)) * scale_factor)
        y = int(((screen_height/2) - (height/1.8)) * scale_factor)

        return f'{width}x{height}+{x}+{y}'