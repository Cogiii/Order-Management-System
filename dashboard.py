from customtkinter import *
from tkcalendar import Calendar
from CTkTable import CTkTable
from datetime import datetime, timedelta
from tkinter import messagebox
from PIL import Image
from resources import Resources
from viewChart import ViewChart
from editOrder import EditOrder

class Dashboard(Resources):
    def __init__(self):
        super().__init__()
        self.table = None
        self.dashboard_view = None

    def show_dashboard(self, master):
        self.dashboard_view = CTkFrame(master=master, fg_color="#ffffff",  width=1180, height=800, corner_radius=0)
        self.dashboard_view.pack_propagate(0)
        self.dashboard_view.pack(side="left")
        
        self.open_file()
        self.top_panel()
        self.metrics_panel()
        self.search_container_panel()
        self.show_table()

    def hide_dashboard(self):
        if self.dashboard_view:
            self.dashboard_view.destroy()
            self.workbook.close()

    def top_panel(self):
        # Title frame or Top Frame
        self.title_frame = CTkFrame(master=self.dashboard_view, fg_color="transparent")
        self.title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 0))
        CTkLabel(master=self.title_frame, text="Orders", font=("Arial Black", 25), text_color=self.COLOR1).pack(anchor="nw", side="left")

    def metrics_panel(self):
        # Metrics frame below the Top Frame
        self.metrics_frame = CTkFrame(master=self.dashboard_view, fg_color="transparent")
        self.metrics_frame.pack(anchor="n", fill="x",  padx=15, pady=(36, 0))

        self.total_orders = self.metrics_label("assets/cart_icon.png", 50, 15, "Total Orders", "0")
        self.total_sales = self.metrics_label("assets/money_icon.png", 60, 10, "Total Sales", f"{self.currency_sign}0")
        self.metric_chart()

    def metrics_label(self, logo, size, pady, title, details):
        # A metrics label template
        self.frame = CTkFrame(master=self.metrics_frame, fg_color=self.COLOR1, width=200, height=80, corner_radius=12)
        self.frame.grid_propagate(0)
        self.frame.pack(side="left", padx=10)

        image_data = Image.open(logo)
        image = CTkImage(dark_image=image_data, light_image=image_data,size=(size, size))
        CTkLabel(master=self.frame, image=image, text="").grid(row=0, column=0, rowspan=2, padx=10, pady=pady)

        CTkLabel(master=self.frame, text=title, text_color=self.COLOR4, font=("Arial Black", 14)).grid(row=0, column=1, sticky="sw")
        labels = CTkLabel(master=self.frame, text=details, text_color="#fff",font=("Arial Black", 18), justify="center")
        labels.grid(row=1, column=1, sticky="nw", pady=(0,10))

        return labels

    def metric_chart(self):
        # Frame for metrics view chart
        self.modification_metric = CTkFrame(master=self.metrics_frame, fg_color="transparent", width=250, height=80, corner_radius=12)
        self.modification_metric.grid_propagate(0)        
        self.modification_metric.pack(side="right")

        CTkLabel(master=self.modification_metric, text="View chart by:", text_color="black", font=("Arial Black", 12)).grid(row=0, column=0, rowspan=2, padx=10, pady=20)

        self.view_chart = CTkComboBox(master=self.modification_metric, button_color=self.COLOR1, border_color=self.COLOR2, button_hover_color=self.COLOR2, dropdown_hover_color=self.COLOR3, width=100, state="readonly", values=["Top Menu"], command=lambda choice: ViewChart(choice, self.start_date, self.end_date, self.sheet))
        self.view_chart.grid(row=0, column=1, rowspan=2, padx=0, pady=10)

    def search_container_panel(self):
        # Search container or Filter container
        self.search_container = CTkFrame(master=self.dashboard_view, height=50, fg_color="#F0F0F0")
        self.search_container.pack(fill="x", pady=(30, 0), padx=27)

        self.search_id_field()

    def search_id_field(self):
         # Input field for Order ID filter
        self.search_order_id = CTkEntry(master=self.search_container, width=300, placeholder_text="Search Order (Press Enter to Search)", border_color=self.COLOR3, border_width=2)
        self.search_order_id.pack(side="left", padx=(13, 10), pady=15)
        self.search_order_id.bind("<Return>", lambda event: self.filter_order_id(event))

        CTkButton(master=self.search_container, text="Edit", text_color=self.COLOR4, font=("Arial Black",14), fg_color=self.COLOR1, hover_color=self.COLOR2, width=100, command=lambda: EditOrder(self.search_order_id.get(), self.EXCEL_FILE, self.workbook, self.sheet, self.menu_price)).pack(side="left", padx=(13, 180), pady=15)

        self.filter_date_field()

    def filter_date_field(self):
        current_date = datetime.now().strftime("%Y-%m-%d") # get current date
        end_date = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d") # add 1 day to current date

        self.start_date = self.date_field(self.search_container, "Start Date:", current_date)
        self.end_date = self.date_field(self.search_container, "End Date:", end_date)

        # Filter Button for start_date and end_date
        CTkButton(master=self.search_container,text="Filter",width=80, text_color=self.COLOR4, font=("Arial Black",14),fg_color=self.COLOR1, hover_color=self.COLOR2, command=self.update_table).pack(side="left", padx=(13,0), pady=15)

    def date_field(self, master, text, date):
        # Date Entry
        CTkLabel(master=master, text=text).pack(side="left", padx=(13, 0), pady=15)
        date_field = CTkEntry(master=self.search_container, width=90, border_color=self.COLOR3, border_width=2)
        date_field.pack(side="left", padx=(13, 0), pady=15)
        date_field.insert(0, date)
        date_field.bind("<1>", lambda event: self.pick_date(date_field, date, event))

        return date_field
    
    def pick_date(self, date_entry, date, event):
        # Display an overidden window from the main window and the calendar
        self.date_window = CTkToplevel()
        self.date_window.grab_set()
        self.date_window.title("Date")
        self.date_window.after(250, lambda: self.date_window.iconbitmap('assets/logo.ico'))
        self.date_window.geometry(self.center_window(self.date_window, 200, 200, self.date_window._get_window_scaling()))
        self.date_window.resizable(0, 0)

        cal = Calendar(self.date_window, selectmode="day", date_pattern="yyyy-mm-dd", day=int(str(date)[-2:]))
        cal.place(x=0, y=0)

        submit_btn = CTkButton(self.date_window, text="submit", fg_color=self.COLOR1, command=lambda: self.grab_date(date_entry, cal, self.date_window))
        submit_btn.place(x=30, y=160)

    def grab_date(self, date_entry, cal, date_window):
        # grabe the date and change the date_field or date_entry
        selected_date = cal.get_date()
        date_entry.delete(0, END)
        date_entry.insert(0, selected_date)
        date_window.destroy()

    def show_table(self):
        # Create table frame 
        self.table_frame = CTkScrollableFrame(master=self.dashboard_view, fg_color="transparent",scrollbar_button_color=self.COLOR2, scrollbar_button_hover_color=self.COLOR3)
        self.table_frame.pack(expand=True, fill="both", padx=10, pady=21)

        self.update_table()

    def update_table(self, table_data=None):
        if table_data is None:
            self.table_data = []
            self.table_data.append(self.table_header) # Append all the values inside the table

            rows_to_append = []
            start_date = int(self.start_date.get().replace("-", ""))
            end_date = int(self.end_date.get().replace("-", ""))

            for row in range(2, self.sheet.max_row + 1):
                date = int(str(self.sheet[f"D{row}"].value).replace("-","")[:8])

                if start_date <= date <= end_date:
                    rows_to_append.append(self.get_data(row))

            self.table_data.extend(reversed(rows_to_append))

            self.total_orders.configure(text=len(self.table_data)-1)
            self.total_sales.configure(text=f"₱{sum([float(bill[4]) for bill in self.table_data[1:]]):.2f}")

        else:
            self.table_data = table_data

        # Create a new table or show "Empty Table" label if no data found
        if self.table:
            self.table.destroy()
        
        self.table = CTkTable(master=self.table_frame, values=self.table_data, colors=["#E6E6E6", "#EEEEEE"], header_color=self.COLOR1, hover_color="#B4B4B4")
        self.table.edit_row(0, text_color="#fff", hover_color=self.COLOR2)
        self.table.pack(expand=True)
        
    
    def get_data(self, row):
        rows_to_append = []
        count_orders = {}
        orders = []

        for col in range(1, len(self.table_header) + 1):
            # just return the value of date instead of date and time
            if isinstance(self.sheet.cell(row=row, column=col).value, datetime):
                rows_to_append.append(self.sheet.cell(row=row, column=col).value.date())
            elif col == 3:
                orders = str(self.sheet.cell(row=row, column=col).value).split(", ")
                for order in orders:
                    if order in count_orders:
                        count_orders[order] += 1
                    else:
                         count_orders[order] = 1
                text_order = ""
                for menu, count in count_orders.items():
                    text_order += f"{count} {menu}, "
                rows_to_append.append(text_order[:-2])
            else:
                 rows_to_append.append(self.sheet.cell(row=row, column=col).value)

        return rows_to_append
    
    def filter_order_id(self, event):
        # check if order_id is a number
        try:
            order_id = int(self.search_order_id.get())
        except:
            messagebox.showerror("Error Message", "Invalid Order ID")
            return
        
        # Get only the row data and call the update table
        for i in range(2, self.sheet.max_row + 1):
            if self.sheet.cell(row=i, column=1).value == order_id:
                filtered_data = [self.table_header, self.get_data(i)]
                self.update_table(filtered_data)
                return
            
        # If there are no data, display error message
        messagebox.showinfo("Error Message", "Order ID entered is not found")
