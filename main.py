from customtkinter import *
from PIL import Image
from tkcalendar import Calendar
from CTkTable import CTkTable
from datetime import datetime, timedelta
from tkinter import messagebox
import matplotlib.pyplot as plt
import openpyxl 


# COLOR PALLETE FROM THIS LINK
# https://colorhunt.co/palette/5356ff378ce767c6e3dff5ff
COLOR1 = "#5356FF"
COLOR2 = "#378CE7"
COLOR3 = "#67C6E3"
COLOR4 = "#DFF5FF"

# Get data from excel
EXCEL_FILE = 'Orders.xlsm'
workbook = openpyxl.load_workbook(EXCEL_FILE, data_only=True) # open worksheet file
sheet = workbook.active # Focus on the first sheet or active sheet


class OrderSystemApp():
    def __init__(self):
        self.app = CTk()
        self.app.geometry(self.center_window(self.app, 850, 700, self.app._get_window_scaling()))
        self.app.resizable(0,0)
        self.app.title("Order System")
        set_appearance_mode("light")
        self.app.iconbitmap("assets/logo.ico")

        self.no_order_id_label = None
        self.setupSidebar() # display side bar
        self.dashboardView() # display dashboard

    # Side Bar FRAME
    def setupSidebar(self):
        # Create side bar frame
        self.sidebar_frame = CTkFrame(master=self.app, fg_color=COLOR1,  width=176, height=700, corner_radius=0)
        self.sidebar_frame.pack_propagate(0)
        self.sidebar_frame.pack(fill="y", anchor="w", side="left")

        # Logo image on the top of the side bar frame
        logo_img_data = Image.open("assets/logo.png") # CHANGE LANG LOGO
        logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(160, 160))
        CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(38, 0), anchor="center")

        # Dashboard Nav Bar
        analytics_img_data = Image.open("assets/analytics_icon.png")
        analytics_img = CTkImage(dark_image=analytics_img_data, light_image=analytics_img_data)
        CTkButton(master=self.sidebar_frame, image=analytics_img, text="Dashboard", fg_color="transparent", font=("Arial Bold", 14), hover_color=COLOR2, anchor="w").pack(anchor="center", ipady=5, pady=(60, 0))

        # Order Nav Bar
        package_img_data = Image.open("assets/cart_icon.png")
        package_img = CTkImage(dark_image=package_img_data, light_image=package_img_data)
        CTkButton(master=self.sidebar_frame, image=package_img, text="Orders", fg_color="transparent", font=("Arial Bold", 14), hover_color=COLOR2, anchor="w").pack(anchor="center", ipady=5, pady=(16, 0))

        # Account Informaion (Initial Only)
        person_image_data = Image.open("assets/person_icon.png")
        person_image = CTkImage(dark_image=person_image_data, light_image=person_image_data)
        CTkButton(master=self.sidebar_frame, image=person_image, text="Admin", fg_color="transparent", font=("Arial Bold", 14), hover_color=COLOR2, anchor="w").pack(anchor="center", ipady=5, pady=(300, 0))
        
        
    # Dashboard View
    def dashboardView(self):
        # Create dashboard frame
        self.dashboard_view = CTkFrame(master=self.app, fg_color="#ffffff",  width=730, height=800, corner_radius=0)
        self.dashboard_view.pack_propagate(0)
        self.dashboard_view.pack(side="left")

        # Title frame or Top Frame
        self.title_frame = CTkFrame(master=self.dashboard_view, fg_color="transparent")
        self.title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        CTkLabel(master=self.title_frame, text="Orders", font=("Arial Black", 25), text_color=COLOR1).pack(anchor="nw", side="left")




        # metrics order frame or status order
        self.metrics_frame = CTkFrame(master=self.dashboard_view, fg_color="transparent")
        self.metrics_frame.pack(anchor="n", fill="x",  padx=15, pady=(36, 0))

        # Total Orders Frame
        self.orders_metric = CTkFrame(master=self.metrics_frame, fg_color=COLOR1, width=200, height=80, corner_radius=12)
        self.orders_metric.grid_propagate(0)
        self.orders_metric.pack(side="left")

        # Image in Total Order Frame
        order_image_data = Image.open("assets/cart_icon.png")
        order_image = CTkImage(dark_image=order_image_data, light_image=order_image_data,size=(50, 50))
        CTkLabel(master=self.orders_metric, image=order_image, text="0").grid(row=0, column=0, rowspan=2, padx=10, pady=15)

        # Total Orders Label
        CTkLabel(master=self.orders_metric, text="Total Orders", text_color=COLOR4, font=("Arial Black", 14)).grid(row=0, column=1, sticky="sw")
        self.total_orders = CTkLabel(master=self.orders_metric, text="", text_color="#fff",font=("Arial Black", 18), justify="center")
        self.total_orders.grid(row=1, column=1, sticky="nw", pady=(0,10))


        # Total Sales metrics Frame
        self.amount_metric = CTkFrame(master=self.metrics_frame, fg_color=COLOR1, width=200, height=80, corner_radius=12)
        self.amount_metric.grid_propagate(0)        
        self.amount_metric.pack(side="left", padx=(15,0))

        # Image in Total Sales Frame
        amount_image_data = Image.open("assets/money_icon.png")
        amount_image = CTkImage(dark_image=amount_image_data, light_image=amount_image_data, size=(60, 60))
        CTkLabel(master=self.amount_metric, image=amount_image, text="").grid(row=0, column=0, rowspan=2, padx=7, pady=10)

        # Total Sales Label
        CTkLabel(master=self.amount_metric, text="Total Sales", text_color=COLOR4, font=("Arial Black", 14)).grid(row=0, column=1, sticky="sw")
        self.total_sales = CTkLabel(master=self.amount_metric, text="₱0", text_color="#fff",font=("Arial Black", 18), justify="center")
        self.total_sales.grid(row=1, column=1, sticky="nw", pady=(0,10))


        # Frame for metrics view chart and edit by order ID using grid
        self.modification_metric = CTkFrame(master=self.metrics_frame, fg_color="transparent", width=250, height=80, corner_radius=12)
        self.modification_metric.grid_propagate(0)        
        self.modification_metric.pack(side="left")

        # View Chart at grid row 0
        CTkLabel(master=self.modification_metric, text="View chart by:", text_color="black", font=("Arial Black", 12)).grid(row=0, column=0, rowspan=1, padx=0, pady=10)
        self.view_chart = CTkComboBox(master=self.modification_metric,
                                      button_color=COLOR1,
                                      border_color= COLOR2,
                                      button_hover_color=COLOR2,
                                      dropdown_hover_color=COLOR3,
                                      width=100,
                                      state="readonly",
                                      values=["Sales", "Top Menu"],
                                      command=self.viewChart)
        self.view_chart.grid(row=0, column=1, rowspan=1, padx=0, pady=10)
        
        # Order Entry at grid row 1
        self.order_id = CTkEntry(master=self.modification_metric, 
                                 placeholder_text="Order ID",
                                 width=100,
                                 height=25,
                                 border_width=2,
                                 border_color=COLOR1,
                                 corner_radius=10)
        self.order_id.grid(row=1,column=0, padx=10)
        CTkButton(master=self.modification_metric, 
                  text="Edit", 
                  text_color=COLOR4, 
                  font=("Arial Black",14),
                  fg_color=COLOR1,
                  hover_color=COLOR2,
                  width=100,
                  command=lambda: self.edit_order(self.order_id.get())).grid(row=1,column=1, padx=5)



        # Search container or Filter container
        self.search_container = CTkFrame(master=self.dashboard_view, height=50, fg_color="#F0F0F0")
        self.search_container.pack(fill="x", pady=(30, 0), padx=27)

        # Input field for Order ID filter
        self.search_order_id = CTkEntry(master=self.search_container, width=150, placeholder_text="Search Order", border_color=COLOR3, border_width=2)
        self.search_order_id.pack(side="left", padx=(13, 0), pady=15)
        self.search_order_id.bind("<Return>", self.filter_order_id)

        current_date = datetime.now().strftime("%Y-%m-%d") # get current date
        end_date = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        # Start Date Entry
        CTkLabel(master=self.search_container, text="Start Date:").pack(side="left", padx=(13, 0), pady=15)
        self.start_date = CTkEntry(master=self.search_container, width=90, border_color=COLOR3, border_width=2)
        self.start_date.pack(side="left", padx=(13, 0), pady=15)
        self.start_date.insert(0, current_date)
        self.start_date.bind("<1>", lambda event: self.pick_date(self.start_date, current_date, event))

        # End Date Entry
        CTkLabel(master=self.search_container, text="End Date:").pack(side="left", padx=(13, 0), pady=15)
        self.end_date = CTkEntry(master=self.search_container, width=90, border_color=COLOR3, border_width=2)
        self.end_date.pack(side="left", padx=(13, 0), pady=15)
        self.end_date.insert(0, end_date)
        self.end_date.bind("<1>", lambda event: (self.pick_date(self.end_date, end_date, event)))

        # Filter Button for start_date and end_date
        CTkButton(master=self.search_container,text="Filter",width=80, text_color=COLOR4, font=("Arial Black",14),fg_color=COLOR1, hover_color=COLOR2, command=self.filter_by_date).pack(side="left", padx=(13,0), pady=15)


        
        self.table_header = ('Order ID','Table Number','Orders','Date','Bill','Status') # Get table header
        self.table_data = [] # Initialize table data

        
        self.table_data.append([header for header in self.table_header]) # Append all the values inside the table
        
        # Iterate over rows in descending order and append the first 12 rows
        rows_to_append = []
        total_sales = 0
        start_date = int(self.start_date.get().replace("-", ""))
        end_date = int(self.end_date.get().replace("-", ""))

        for row in range(2, 13):
            values = []
            date = int(str(sheet[f"D{row}"].value).replace("-","")[:8])
            
            if start_date <= date <= end_date:
                for col in range(1, len(self.table_header) + 1):
                    if isinstance(sheet.cell(row=row, column=col).value, datetime):
                        values.append(sheet.cell(row=row, column=col).value.date())

                    else:
                        values.append(sheet.cell(row=row, column=col).value)
                        
                        if col == 5:
                            total_sales += float(sheet.cell(row=row,column=col).value)

                rows_to_append.append(values)

        self.table_data.extend(rows_to_append)
        
        self.total_orders.configure(text=len(self.table_data)-1)
        self.total_sales.configure(text=f"₱{total_sales:.2f}")

        # Create table frame 
        self.table_frame = CTkScrollableFrame(master=self.dashboard_view, fg_color="transparent",scrollbar_button_color=COLOR2, scrollbar_button_hover_color=COLOR3)
        self.table_frame.pack(expand=True, fill="both", padx=10, pady=21)
        
        # Check if table data has values
        if len(self.table_data) > 0:
            # Show tables from the table data
            self.table = CTkTable(master=self.table_frame, values=self.table_data, colors=["#E6E6E6", "#EEEEEE"], header_color=COLOR1, hover_color="#B4B4B4")
            self.table.edit_row(0, text_color="#fff", hover_color=COLOR2)
            self.table.pack(expand=True)

    # For edit order window
    def edit_order(self, order_id):
        # Create another window that override the order system app
        self.edit_window = CTkToplevel()
        self.edit_window.grab_set()
        self.edit_window.title("Edit")
        self.edit_window.after(250, lambda: self.edit_window.iconbitmap('assets/logo.ico'))
        self.edit_window.geometry(self.center_window(self.edit_window, 500, 350, self.edit_window._get_window_scaling()))
        self.edit_window.resizable(0, 0)

        # Display Order ID
        CTkLabel(master=self.edit_window, text=f"Order ID: {order_id}", text_color=COLOR1,font=("Arial Black", 24), justify="center").place(x=10,y=10)

        # Display table number by Option
        CTkLabel(master=self.edit_window, text="Table Number:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=60)
        self.table_number = CTkComboBox(master=self.edit_window,
                                        button_color=COLOR1,
                                        border_color= COLOR2,
                                        button_hover_color=COLOR2,
                                        width=100,
                                        dropdown_hover_color=COLOR3,
                                        state="readonly",
                                        values=['1','2','3','4','5','6','7','8','9','10','TakeOut'],
                                        )
        self.table_number.place(x=150,y=60)

        # Display Orders
        CTkLabel(master=self.edit_window, text="Orders:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=100)
        
        def slider_event(value, label):
            label.configure(text=int(value))

        def orders_check(slider, check_var):
            if check_var.get() == 'off':
                slider.configure(state='disabled')
            else:
                slider.configure(state='normal')

        pizza_var = StringVar(value="on")
        carbonara_var = StringVar(value="on")
        chicken_var = StringVar(value="on")
        pancit_var = StringVar(value="on")
        lasagna_var = StringVar(value="on")

        # Pizza Check Box
        self.pizza_checkbox = CTkCheckBox(self.edit_window, text="Pizza", width=5, command= lambda: orders_check(self.pizza_slider, pizza_var),variable=pizza_var, onvalue="on", offvalue="off")
        self.pizza_checkbox.place(x=10, y=130)
        self.pizza_quantity = CTkLabel(master=self.edit_window, text="0", text_color="black",font=("Arial", 15), justify="center")
        self.pizza_quantity.place(x=105,y=129)
        self.pizza_slider = CTkSlider(self.edit_window, from_=0, to=10, command=lambda value: slider_event(value,self.pizza_quantity), width=100, state='disabled'if pizza_var == 'off'else 'normal')
        self.pizza_slider.place(x=125, y=135)

        # Carbonara Check Box
        self.carbonara_checkbox = CTkCheckBox(self.edit_window, text="Carbonara", width=5, command= lambda: orders_check(self.carbonara_slider, carbonara_var),variable=carbonara_var, onvalue="on", offvalue="off")
        self.carbonara_checkbox.place(x=10, y=170)
        self.carbonara_quantity = CTkLabel(master=self.edit_window, text="0", text_color="black",font=("Arial", 15), justify="center")
        self.carbonara_quantity.place(x=105,y=169)
        self.carbonara_slider = CTkSlider(self.edit_window, from_=0, to=10, command=lambda value: slider_event(value,self.carbonara_quantity), width=100, state='disabled'if carbonara_var == 'off'else 'normal')
        self.carbonara_slider.place(x=125, y=175)

        # Chicken Check Box
        self.chicken_checkbox = CTkCheckBox(self.edit_window, text="Chicken", width=5, command= lambda: orders_check(self.chicken_slider, chicken_var), variable=chicken_var, onvalue="on", offvalue="off")
        self.chicken_checkbox.place(x=10, y=210)
        self.chicken_quantity = CTkLabel(master=self.edit_window, text="0", text_color="black",font=("Arial", 15), justify="center")
        self.chicken_quantity.place(x=105,y=209)
        self.chicken_slider = CTkSlider(self.edit_window, from_=0, to=10, command=lambda value: slider_event(value,self.chicken_quantity), width=100, state='disabled'if chicken_var == 'off'else 'normal')
        self.chicken_slider.place(x=125, y=215)

        # Pancit Check Box
        self.pancit_checkbox = CTkCheckBox(self.edit_window, text="Pancit", width=5,command= lambda: orders_check(self.pancit_slider, pancit_var), variable=pancit_var, onvalue="on", offvalue="off")
        self.pancit_checkbox.place(x=230, y=130)
        self.pancit_quantity = CTkLabel(master=self.edit_window, text="0", text_color="black",font=("Arial", 15), justify="center")
        self.pancit_quantity.place(x=325,y=129)
        self.pancit_slider = CTkSlider(self.edit_window, from_=0, to=10, command=lambda value: slider_event(value,self.pancit_quantity), width=100, state='disabled'if pancit_var == 'off'else 'normal')
        self.pancit_slider.place(x=345, y=135)

        # Lasagna Check Box
        self.lasagna_checkbox = CTkCheckBox(self.edit_window, text="Lasagna", width=5,command= lambda: orders_check(self.lasagna_slider, lasagna_var), variable=lasagna_var, onvalue="on", offvalue="off")
        self.lasagna_checkbox.place(x=230, y=170)
        self.lasagna_quantity = CTkLabel(master=self.edit_window, text="0", text_color="black",font=("Arial", 15), justify="center")
        self.lasagna_quantity.place(x=325,y=169)
        self.lasagna_slider = CTkSlider(self.edit_window, from_=0, to=10, command=lambda value: slider_event(value,self.lasagna_quantity), width=100, state='disabled'if lasagna_var == 'off'else 'normal')
        self.lasagna_slider.place(x=345, y=175)

        # Total Bill Entry
        CTkLabel(master=self.edit_window, text="Total Bill:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=250)
        self.total_bill = CTkEntry(master=self.edit_window, placeholder_text="₱0",text_color="black",font=("Arial", 15), border_color=COLOR4,height=10)
        self.total_bill.place(x=80, y=253)

        # Amount Entry
        CTkLabel(master=self.edit_window, text="Amount:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=280)
        self.total_bill = CTkEntry(master=self.edit_window, placeholder_text="₱0",text_color="black",font=("Arial", 15), border_color=COLOR4,height=10)
        self.total_bill.place(x=80, y=283)

        # Amount Entry
        CTkLabel(master=self.edit_window, text="Change:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=310)
        self.total_change = CTkLabel(master=self.edit_window, text="₱0.00",text_color="black",font=("Arial", 15))
        self.total_change.place(x=80, y=313)

        CTkLabel(master=self.edit_window, text="Status:", text_color="black",font=("Arial", 15), justify="center").place(x=265,y=253)
        self.status = CTkComboBox(master=self.edit_window,
                                        button_color=COLOR1,
                                        border_color= COLOR2,
                                        button_hover_color=COLOR2,
                                        width=100,
                                        dropdown_hover_color=COLOR3,
                                        state="readonly",
                                        values=['Serving','Pending','Dining','Completed'],
                                        )
        self.status.place(x=325,y=253)


    # A function that view the charts 
    def viewChart(self, choice):
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        filtered_start_date = int(self.start_date.get().replace("-", ""))
        filtered_end_date = int(self.end_date.get().replace("-", ""))

        # Error if start date is ahead from the end date
        if filtered_start_date > filtered_end_date:
            messagebox.showerror("Error Message", "Invalid Date")
        else:
            # If user selects menu use Pie chart to compare the top menu
            if choice == 'Top Menu':
                total_orders = {}

                # Get the row that validate the dates
                for row in range(2, sheet.max_row + 1):
                    date = int(str(sheet[f"D{row}"].value).replace("-","")[:8])

                    # Get the orders
                    if filtered_start_date <= date <= filtered_end_date:
                        orders = str(sheet[f"C{row}"].value).split(", ")

                        # Count how many orders and store in the list or dictionary
                        for order in orders:
                            if order in total_orders:
                                total_orders[order] += 1
                            else:
                                total_orders[order] = 1

                menu = [menu for menu in total_orders.keys()]
                values = [value for value in total_orders.values()]
                
                if len(menu) or len(values) > 0:
                    # Create Pie Chart
                    wp = {'linewidth': 1, 'edgecolor': COLOR1}
                    explode = [0.05] * len(values)
                    fig, ax = plt.subplots(figsize=(10, 7))
                    wedges, texts, autotexts = ax.pie(values,
                                                    autopct='%1.1f%%',
                                                    explode=explode,
                                                    labels=menu,
                                                    shadow=True,
                                                    startangle=90,
                                                    wedgeprops=wp,
                                                    textprops=dict(color="black"))
                    # Adding legend
                    ax.legend(wedges, menu,
                            title="Cars",
                            loc="center left",
                            bbox_to_anchor=(1, 0, 0.5, 1))
                    
                    plt.setp(autotexts, size=8, weight="bold")
                    plt.title(f"Top Menu ({start_date} - {end_date})")
                    plt.show()
                else:
                    messagebox.showerror("Error Message", "No Data Found between those dates!!!")
            else:
                pass
                # FOR SAAAALLLEEEESSSSS

    def filter_by_date(self):
        start_date = int(self.start_date.get().replace("-", ""))
        end_date = int(self.end_date.get().replace("-", ""))
        filtered_data = []
        total_sales = 0

        filtered_data.append(self.table_header)
        if start_date > end_date:
            messagebox.showerror("Error Message", "Invalid Date")
        else:
            rows_to_append = []
            for row in range(2, sheet.max_row + 1):
                date = int(str(sheet[f"D{row}"].value).replace("-","")[:8])
                row_data = []
                
                if start_date <= date <= end_date:
                    for col in range(1, len(self.table_header) + 1):
                        if isinstance(sheet.cell(row=row, column=col).value, datetime):
                            row_data.append(sheet.cell(row=row, column=col).value.date())
                        else:
                            row_data.append(sheet.cell(row=row, column=col).value)

                            if col == 5:
                                total_sales += float(sheet.cell(row=row,column=col).value)
                    
                    rows_to_append.append(row_data)
            filtered_data.extend(rows_to_append)

            # Create a new table or show "Empty Table" label if no data found
            if len(filtered_data) > 0:  # Check if there are rows other than header
                self.table.destroy()
                self.total_orders.configure(text=len(filtered_data)-1)
                self.total_sales.configure(text=f"₱{total_sales:.2f}")

                self.table = CTkTable(master=self.table_frame, values=filtered_data, colors=["#E6E6E6", "#EEEEEE"], header_color=COLOR1, hover_color="#B4B4B4")
                self.table.edit_row(0, text_color="#fff", hover_color=COLOR2)
                self.table.pack(expand=True)


    def filter_order_id(self, event):
        order_id = int(self.search_order_id.get())
        filtered_data = []
        row_order_ID = None

        for i in range(2, sheet.max_row + 1):
            if sheet.cell(row=i, column=1).value == order_id:
                row_order_ID = i

        # check of order ID Found
        if row_order_ID:
            # Append table header
            filtered_data.append(self.table_header)

            # just return the value of date instead of date and time
            row_data = [ sheet.cell(row=row_order_ID, column=col).value.date()
                        if isinstance(sheet.cell(row=row_order_ID, column=col).value, datetime) 
                        else sheet.cell(row=row_order_ID, column=col).value
                        for col in range(1, len(self.table_header) + 1)]
                        
            if order_id == row_data[0]:  # Check if order_id is equal to the first column (Order ID)
                filtered_data.append(row_data)
                
            # Create a new table or show "Empty Table" label if no data found
            # Check if there are rows other than header
            self.table.destroy()

            self.table = CTkTable(master=self.table_frame, values=filtered_data, colors=["#E6E6E6", "#EEEEEE"], header_color=COLOR1, hover_color="#B4B4B4")
            self.table.edit_row(0, text_color="#fff", hover_color=COLOR2)
            self.table.pack(expand=True)
        else:
            messagebox.showerror("Error Message", "Invalid Order ID")


    def pick_date(self, date_entry, date,event):
        self.date_window = CTkToplevel()
        self.date_window.grab_set()
        self.date_window.title("Date")
        self.date_window.after(250, lambda: self.date_window.iconbitmap('assets/logo.ico'))
        self.date_window.geometry(self.center_window(self.date_window, 200, 200, self.date_window._get_window_scaling()))
        self.date_window.resizable(0, 0)

        cal = Calendar(self.date_window, selectmode="day", date_pattern="yyyy-mm-dd", day=int(str(date)[-2:]))
        cal.place(x=0, y=0)

        submit_btn = CTkButton(self.date_window, text="submit", fg_color=COLOR1, command=lambda: self.grab_date(date_entry, cal, self.date_window))
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
