from customtkinter import *
from PIL import Image
from tkcalendar import Calendar
from datetime import datetime
from CTkTable import CTkTable
import pandas as pd
import openpyxl 

color1 = "#5356FF"
color2 = "#378CE7"
color3 = "#67C6E3"
color4 = "#DFF5FF"

class OrderSystemApp():
    def __init__(self):
        self.app = CTk()
        self.app.geometry(self.center_window(self.app, 850, 700, self.app._get_window_scaling()))
        self.app.resizable(0,0)
        self.app.title("Order System")
        set_appearance_mode("light")
        
        self.setup_sidebar()
        self.dashboard_view()

    # Side Bar FRAME
    def setup_sidebar(self):
        # Create side bar frame
        self.sidebar_frame = CTkFrame(master=self.app, fg_color=color1,  width=176, height=700, corner_radius=0)
        self.sidebar_frame.pack_propagate(0)
        self.sidebar_frame.pack(fill="y", anchor="w", side="left")

        # Logo image on the top of the side bar frame
        logo_img_data = Image.open("assets/logo.png") # CHANGE LANG LOGO
        logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(160, 160))
        CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(38, 0), anchor="center")

        # Dashboard Page
        analytics_img_data = Image.open("assets/analytics_icon.png")
        analytics_img = CTkImage(dark_image=analytics_img_data, light_image=analytics_img_data)
        CTkButton(master=self.sidebar_frame, image=analytics_img, text="Dashboard", fg_color="transparent", font=("Arial Bold", 14), hover_color=color2, anchor="w").pack(anchor="center", ipady=5, pady=(60, 0))

        package_img_data = Image.open("assets/cart_icon.png")
        package_img = CTkImage(dark_image=package_img_data, light_image=package_img_data)
        CTkButton(master=self.sidebar_frame, image=package_img, text="Orders", fg_color="transparent", font=("Arial Bold", 14), hover_color=color2, anchor="w").pack(anchor="center", ipady=5, pady=(16, 0))

        # Accoung Informaion (Initial Only)
        person_image_data = Image.open("assets/person_icon.png")
        person_image = CTkImage(dark_image=person_image_data, light_image=person_image_data)
        CTkButton(master=self.sidebar_frame, image=person_image, text="Admin", fg_color="transparent", font=("Arial Bold", 14), hover_color=color2, anchor="w").pack(anchor="center", ipady=5, pady=(300, 0))
        
    def dashboard_view(self):
        # Create dashboard frame
        self.dashboard_view = CTkFrame(master=self.app, fg_color="#ffffff",  width=730, height=800, corner_radius=0)
        self.dashboard_view.pack_propagate(0)
        self.dashboard_view.pack(side="left")

        # Title frame or Top Frame
        title_frame = CTkFrame(master=self.dashboard_view, fg_color="transparent")
        title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        CTkLabel(master=title_frame, text="Orders", font=("Arial Black", 25), text_color=color1).pack(anchor="nw", side="left")



        # metrics order frame or status order
        metrics_frame = CTkFrame(master=self.dashboard_view, fg_color="transparent")
        metrics_frame.pack(anchor="n", fill="x",  padx=27, pady=(36, 0))

        # Order metrics frame
        orders_metric = CTkFrame(master=metrics_frame, fg_color=color1, width=200, height=80, corner_radius=12) # create frame
        orders_metric.grid_propagate(0)
        orders_metric.pack(side="left")
        order_image_data = Image.open("assets/cart_icon.png") # import png
        order_image = CTkImage(dark_image=order_image_data, light_image=order_image_data,size=(50, 50))
        CTkLabel(master=orders_metric, image=order_image, text="").grid(row=0, column=0, rowspan=2, padx=10, pady=15)
        # Order metrix labels
        CTkLabel(master=orders_metric, text="Orders", text_color=color4, font=("Arial Black", 14)).grid(row=0, column=1, sticky="sw")
        CTkLabel(master=orders_metric, text="123", text_color="#fff",font=("Arial Black", 18), justify="center").grid(row=1, column=1, sticky="nw", pady=(0,10))

        # Amoung metrics frame
        amount_metric = CTkFrame(master=metrics_frame, fg_color=color1, width=200, height=80, corner_radius=12)
        amount_metric.grid_propagate(0)        
        amount_metric.pack(side="left", padx=27)
        amount_image_data = Image.open("assets/money_icon.png")
        amount_image = CTkImage(dark_image=amount_image_data, light_image=amount_image_data, size=(60, 60))
        CTkLabel(master=amount_metric, image=amount_image, text="").grid(row=0, column=0, rowspan=2, padx=7, pady=10)
        CTkLabel(master=amount_metric, text="Total Amount", text_color=color4, font=("Arial Black", 14)).grid(row=0, column=1, sticky="sw")
        CTkLabel(master=amount_metric, text="₱1300", text_color="#fff",font=("Arial Black", 18), justify="center").grid(row=1, column=1, sticky="nw", pady=(0,10))



        # Search container or Filter container
        search_container = CTkFrame(master=self.dashboard_view, height=50, fg_color="#F0F0F0")
        search_container.pack(fill="x", pady=(30, 0), padx=27)

        # Input field for Order ID filter
        CTkEntry(master=search_container, width=250, placeholder_text="Search Order", border_color=color3, border_width=2).pack(side="left", padx=(13, 0), pady=15)

        current_date = datetime.now().strftime("%Y-%m-%d")
        # Start Date Entry
        CTkLabel(master=search_container, text="Start Date:").pack(side="left", padx=(13, 0), pady=15)
        self.start_date = CTkEntry(master=search_container, width=90, border_color=color3, border_width=2)
        self.start_date.pack(side="left", padx=(13, 0), pady=15)
        self.start_date.insert(0, current_date)
        self.start_date.bind("<1>", lambda event: self.pick_date(self.start_date, event))

        # End Date Entry
        CTkLabel(master=search_container, text="End Date:").pack(side="left", padx=(13, 0), pady=15)
        self.end_date = CTkEntry(master=search_container, width=90, border_color=color3, border_width=2)
        self.end_date.pack(side="left", padx=(13, 0), pady=15)
        self.end_date.insert(0, current_date)
        self.end_date.bind("<1>", lambda event: self.pick_date(self.end_date, event))


        
        # Get data from excel
        EXCEL_FILE = 'Orders.xlsx'
        workbook = openpyxl.load_workbook(EXCEL_FILE)
        sheet = workbook.active
        table_header = [sheet.cell(row=1, column=i).value for i in range(1, sheet.max_column+1)]
        table_data = []

        table_data.append([header for header in table_header])
        rows_to_append = []

        for row in range(2, sheet.max_row + 1):
            values = [
                sheet.cell(row=row, column=col).value.date() if isinstance(sheet.cell(row=row, column=col).value, datetime) else sheet.cell(row=row, column=col).value
                for col in range(1, len(table_header) + 1)
            ]
            rows_to_append.append(values)

        table_data.extend(rows_to_append)

        # Create table frame 
        table_frame = CTkScrollableFrame(master=self.dashboard_view, fg_color="transparent",)
        table_frame.pack(expand=True, fill="both", padx=10, pady=21)
        if len(table_data) > 0:
            # Show tables from the table data
            table = CTkTable(master=table_frame, values=table_data, colors=["#E6E6E6", "#EEEEEE"], header_color=color1, hover_color="#B4B4B4")
            table.edit_row(0, text_color="#fff", hover_color=color2)
            table.pack(expand=True)

        else:
            CTkLabel(master=table_frame, text="Empty Table", text_color="#848787",font=("Arial Black", 24), justify="center").pack(anchor="center")


    def pick_date(self, date_entry, event):
        self.date_window = CTkToplevel()
        self.date_window.grab_set()
        self.date_window.title("Date")
        self.date_window.geometry(self.center_window(self.date_window, 200, 200, self.date_window._get_window_scaling()))
        self.date_window.resizable(0,0)
        cal = Calendar(self.date_window, selectmode="day", date_pattern="yyyy-mm-dd")
        cal.place(x=0, y=0)

        submit_btn = CTkButton(self.date_window, text="submit", fg_color=color1, command=lambda: self.grab_date(date_entry, cal, self.date_window))
        submit_btn.place(x=30, y=160)

    def grab_date(self, date_entry, cal, date_window):
        selected_date = cal.get_date()
        date_entry.delete(0, END)
        date_entry.insert(0, selected_date)
        date_window.destroy()

    def center_window(self, Screen: CTk, width: int, height: int, scale_factor: float = 1.0):
        screen_width = self.app.winfo_screenwidth()
        screen_height = self.app.winfo_screenheight()
        
        x = int(((screen_width/2) - (width/2)) * scale_factor)
        y = int(((screen_height/2) - (height/2)) * scale_factor)

        return f'{width}x{height}+{x}+{y}'
    
    def run(self):

        self.app.mainloop()


if __name__ == "__main__":
    order_system_app = OrderSystemApp()
    order_system_app.run()
