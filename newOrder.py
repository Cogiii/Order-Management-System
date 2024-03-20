from customtkinter import *
from tkcalendar import Calendar
from datetime import datetime, timedelta
from tkinter import messagebox
from resources import Resources

class NewOrder(Resources):
    def __init__(self):
        super().__init__()
        self.order_view = None

    def show_new_order(self, master):
        self.master = master

        self.order_view = CTkFrame(master=self.master, fg_color="#ffffff",  width=1180, height=800, corner_radius=0)
        self.order_view.pack_propagate(0)
        self.order_view.pack(side="left")

        self.open_file()
        self.top_panel()
        self.order_panel()

    def hide_new_order(self):
        if self.order_view:
            self.order_view.destroy()
            self.workbook.close()

    def top_panel(self):
        # Title frame for order_view
        self.title_frame = CTkFrame(master=self.order_view, fg_color="transparent")
        self.title_frame.pack(anchor="n", fill="x",  padx=27, pady=(29, 35))

        max_id = max([self.sheet[f"A{row}"].value for row in range(2, self.sheet.max_row+1)])
        self.next_order_id = max_id+1

        CTkLabel(master=self.title_frame, text="Order Number: ", font=("Arial Black", 30), text_color=self.COLOR1).pack(anchor="nw", side="left")
        self.Order_ID = CTkLabel(master=self.title_frame, text=f"{self.next_order_id}", text_color=self.COLOR1, font=("Arial Black", 30))
        self.Order_ID.pack(anchor="nw", side="left")

    def order_panel(self):
        # Create order frame
        self.order_frame = CTkFrame(master=self.order_view, fg_color="#d2d3d6")
        self.order_frame.pack(anchor="n", fill="x",  padx=27, pady=(0, 0))


        # Table Number Field
        self.table_number_field = CTkFrame(master=self.order_frame, fg_color="transparent")
        self.table_number_field.pack(anchor="n", fill="x",  padx=27, pady=(20,0))
        
        table_num_options = ['1','2','3','4','5','6','7','8','9','10','TakeOut']
        CTkLabel(master=self.table_number_field, text="Table Number:", font=("Arial Black", 18)).pack(side="left")

        self.table_number = CTkComboBox(master=self.table_number_field, button_color=self.COLOR1, border_color=self.COLOR2, button_hover_color=self.COLOR2, width=150, dropdown_hover_color=self.COLOR3, state="readonly", values=table_num_options)
        self.table_number.pack(side="left", padx=20)

        self.menu_orders_field()

    def menu_orders_field(self):
        # Orders Field
        self.orders_field = CTkFrame(master=self.order_frame, fg_color="transparent")
        self.orders_field.grid_propagate(0)
        self.orders_field.pack(anchor="n", fill="x", padx=27, pady=(30,0))
        
        CTkLabel(master=self.orders_field, text="Orders:", font=("Arial Black", 20)).grid(row=0, column=3, pady=(0,5))

        self.row_data_orders ={"Pizza": 0,
                         "Carbonara": 0,
                         "Chicken": 0,
                         "Pancit": 0,
                         "Lasagna": 0}
        
        # create component for each menu
        self.pizza_checkbox, self.pizza_quantity, self.pizza_slider = self.create_components("Pizza", 1, 0)
        self.carbonara_checkbox, self.carbonara_quantity, self.carbonara_slider = self.create_components("Carbonara", 2, 0)
        self.chicken_checkbox, self.chicken_quantity, self.chicken_slider = self.create_components("Chicken", 3, 0)
        self.pancit_checkbox, self.pancit_quantity, self.pancit_slider = self.create_components("Pancit", 1, 4)
        self.lasagna_checkbox, self.lasagna_quantity, self.lasagna_slider = self.create_components("Lasagna", 2, 4)

        # add order button
        self.add_order_btn = CTkButton(self.order_view, text="Add Order", font=("Arial", 15), fg_color=self.COLOR1, hover_color=self.COLOR2, text_color=self.COLOR4,width=150, height=35, command=lambda: self.order_information_window(self.next_order_id, self.table_number.get(), self.row_data_orders))
        self.add_order_btn.pack(anchor="ne", padx=30, pady=20)

    def create_components(self, item_name, x, y):
        # Create widgets
        checkbox = CTkCheckBox(master=self.orders_field, text=item_name, font=("Arial", 15), width=10, onvalue="on", offvalue="off", command=lambda: self.orders_check(slider, quantity_label, checkbox, self.row_data_orders, item_name))
        checkbox.grid(row=x, column=y, padx=(0, 30), pady=10, sticky="w")

        quantity_label = CTkLabel(master=self.orders_field, text="0", font=("Arial", 20), justify="center", width=30)
        quantity_label.grid(row=x, column=y + 1, padx=(0, 0), pady=10)

        slider = CTkSlider(master=self.orders_field, from_=0, to=10, width=150, state='disabled', command=lambda value: self.slider_event(value, quantity_label, item_name, self.row_data_orders))
        slider.grid(row=x, column=y + 2, padx=(0, 0), pady=10)

        return checkbox, quantity_label, slider
    
    def order_information_window(self, order_id, table_number, orders):
        # check if any of the order is not 0
        num = 0
        for val in orders.values():
            if val > num:
                num = val
                break
        
        # display error when either of table number or orders is empty
        if table_number == "" or num <= 0:
            messagebox.showerror("Error Message", "Invalid Order!")
            return

        date = datetime.now().strftime("%Y-%m-%d") # Date today
        bill = sum(orders[order]*self.menu_price[order] for order in orders) # sum of all the orders
        status = "Serving"

        # create the window
        order_information_window = CTkToplevel()
        order_information_window.grab_set()
        order_information_window.title("Confirmation")
        order_information_window.after(250, lambda: order_information_window.iconbitmap('assets/logo.ico'))
        order_information_window.geometry(self.center_window(order_information_window, 260, 300, order_information_window._get_window_scaling()))
        order_information_window.resizable(0, 0)

        # Create Label to display the information
        CTkLabel(master=order_information_window, text=f"Order ID\t\t: {order_id}", font=("Arial", 15)).place(x=10, y=10)
        CTkLabel(master=order_information_window, text=f"Table Number\t: {table_number}", font=("Arial", 15)).place(x=10, y=30)
        CTkLabel(master=order_information_window, text=f"Date\t\t: {date}", font=("Arial", 15)).place(x=10, y=50)
        CTkLabel(master=order_information_window, text=f"Bill\t\t: ₱{bill}", font=("Arial", 15)).place(x=10, y=70)
        CTkLabel(master=order_information_window, text=f"Status\t\t: {status}", font=("Arial", 15)).place(x=10, y=90)
        CTkLabel(master=order_information_window, text="Orders:", font=("Arial", 15)).place(x=10, y=120)

        # Displaying orders
        y_offset = 120  # Initial y-coordinate for the first order
        for menu, count in orders.items():
            if count > 0:
                order_text = f"\u2022 {count} {menu}"  # Using bullet character \u2022
                CTkLabel(order_information_window, text=order_text, font=("Arial", 13)).place(x=90, y=y_offset)
                y_offset += 20  # Increment y-coordinate for next order

        # Confirmation Button
        CTkButton(master=order_information_window, text="Confirm", fg_color=self.COLOR1, font=("Arial", 15), command=lambda: self.add(order_id, table_number, orders, date, bill, status, order_information_window)).place(x=60, y=260)
    
    def add(self, order_id, table_num, orders, date, bill, status, window):
        order = ""
        # get the menu separated by comma to it in excel
        for menu, num in orders.items():
            order += (menu + ", ") * num
        order = order[:-2]
        row = (order_id, table_num if table_num == "TakeOut" else int(table_num), order, date, bill, status)
        self.sheet.append(row)
        self.workbook.save(self.EXCEL_FILE)

        window.destroy()
        self.order_view.destroy()
        self.show_new_order(self.master)

    # just get the value of sliders and display it dynamically
    def slider_event(self, value, label, item_name, row_data_orders, amount_entry=None):
        label.configure(text=int(value))
        row_data_orders[item_name] = int(value)

        bill = sum(row_data_orders[order]*self.menu_price[order] for order in row_data_orders)
        if amount_entry:
            amount_entry.delete(0, END)
            amount_entry.insert(0, f"{bill:.2f}")

    # for slider event set disabled if checkbox is off
    def orders_check(self, slider, quantity_label, check_var, row_data_orders, item_name, amount_entry=None):
        slider.configure(state='disabled' if check_var.get() == 'off' else 'normal')
        if check_var.get() == "off":
            quantity_label.configure(text=0)
            row_data_orders[item_name] = 0

        bill = sum(row_data_orders[order]*self.menu_price[order] for order in row_data_orders)
        if amount_entry:
            amount_entry.delete(0, END)
            amount_entry.insert(0, f"{bill:.2f}")