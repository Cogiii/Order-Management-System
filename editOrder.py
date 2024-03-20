from customtkinter import *
from resources import Resources
from tkinter import messagebox
from datetime import datetime, timedelta

class EditOrder(Resources):
    def __init__(self, order_id, excel_file, workbook, sheet, menu_price):
        super().__init__()
        self.EXCEL_FILE = exce
        self.workbook = workbook
        self.sheet = sheet
        self.menu_price = menu_price

        self.get_order_row(order_id)

    # For edit order window
    def get_order_row(self, order_id):
        try:
            order_id = int(order_id)
        except:
            messagebox.showerror("Error Message", "Invalid Order ID")
            return
        
        self.row = 0

        # get order id rows
        row_data = []
        for row in range(2,self.sheet.max_row + 1):
            if int(self.sheet[f'A{row}'].value) == int(order_id):
                self.row = row

                # 1 to 6 because there only 6 column on the table
                for col in range(1, 7):
                    if isinstance(self.sheet.cell(row=row, column=col).value, datetime):
                        row_data.append(self.sheet.cell(row=row, column=col).value.date())

                    elif col == 3:
                        orders = str(self.sheet.cell(row=row, column=col).value).split(", ")
                        row_data.append(orders)

                    else:
                        row_data.append(self.sheet.cell(row=row, column=col).value)
                    
                break        
        
        # Prompt error when row_data is empty
        if len(row_data) <= 0:
            messagebox.showerror("Error Message", "Invalid Order ID")
            return
        
        self.edit_order_window(row_data, order_id)

    def edit_order_window(self, row_data, order_id):

        # Create another window that override the order system app
        self.edit_window = CTkToplevel()
        self.edit_window.grab_set()
        self.edit_window.title("Edit")
        self.edit_window.after(250, lambda: self.edit_window.iconbitmap('assets/logo.ico'))
        self.edit_window.geometry(self.center_window(self.edit_window, 500, 350, self.edit_window._get_window_scaling()))
        self.edit_window.resizable(0, 0)

        # Display Order ID
        CTkLabel(master=self.edit_window, text=f"Order ID: {row_data[0]}", text_color=self.COLOR1,font=("Arial Black", 24), justify="center").place(x=10,y=10)


        # Display table number by Option
        CTkLabel(master=self.edit_window, text="Table Number:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=60)

        table_num = StringVar(value=f"{row_data[1]}")
        self.table_number = CTkComboBox(master=self.edit_window, button_color=self.COLOR1, border_color=self.COLOR2, button_hover_color=self.COLOR2, width=100, dropdown_hover_color=self.COLOR3, state="readonly", variable=table_num, 
                                            values=['1','2','3','4','5','6','7','8','9','10','TakeOut'])
        self.table_number.place(x=150,y=60)

        # Display Orders
        CTkLabel(master=self.edit_window, text="Orders:", text_color="black",font=("Arial", 15), justify="center").place(x=10,y=100)

        # Check Box Variables
        pizza_var = StringVar(value="on" if "Pizza" in row_data[2] else "off")
        carbonara_var = StringVar(value="on" if "Carbonara" in row_data[2] else "off")
        chicken_var = StringVar(value="on" if "Chicken" in row_data[2] else "off")
        pancit_var = StringVar(value="on" if "Pancit" in row_data[2] else "off")
        lasagna_var = StringVar(value="on" if "Lasagna" in row_data[2] else "off")

        # Get and count the orders
        self.row_data_orders = {"Pizza": 0, 
                                "Carbonara": 0,
                                "Chicken": 0,
                                "Pancit": 0,
                                "Lasagna": 0}
        for order in row_data[2]:
            self.row_data_orders[order] += 1

        self.pizza_checkbox, self.pizza_quantity, self.pizza_slider = self.create_components("Pizza", 10, 130, pizza_var)
        self.carbonara_checkbox, self.carbonara_quantity, self.carbonara_slider = self.create_components("Carbonara", 10, 170, carbonara_var)
        self.chicken_checkbox, self.chicken_quantity, self.chicken_slider = self.create_components("Chicken", 10, 210, chicken_var)
        self.pancit_checkbox, self.pancit_quantity, self.pancit_slider = self.create_components("Pancit", 230, 130, pancit_var)
        self.lasagna_checkbox, self.lasagna_quantity, self.lasagna_slider = self.create_components("Lasagna", 230, 170, lasagna_var)

# ----------------------------------------------------------------------------------------------------------------------------------- #

        # Total Bill Entry
        CTkLabel(master=self.edit_window, text="Total Bill:", text_color="black",font=("Arial", 14), justify="center").place(x=10,y=250)

        self.total_bill = CTkEntry(master=self.edit_window, placeholder_text="0.00",text_color="black",font=("Arial", 13), border_color=self.COLOR3, height=10, fg_color="transparent")

        self.total_bill.insert(0, f"{row_data[4]:.2f}")
        self.total_bill.place(x=100, y=253)

# ----------------------------------------------------------------------------------------------------------------------------------- #

        # Amount Entry
        amount_pay = StringVar(value="0.00")
        CTkLabel(master=self.edit_window, text="Amount Pay:", text_color="black",font=("Arial", 14), justify="center").place(x=10,y=280)

        self.amount_pay = CTkEntry(master=self.edit_window, placeholder_text="0",text_color="black",font=("Arial", 13), border_color=self.COLOR3,height=10, fg_color="transparent", textvariable=amount_pay)
        self.amount_pay.place(x=100, y=283)
        self.amount_pay.bind("<Return>", lambda event: self.display_change(event, self.amount_pay.get(), self.total_bill.get()))

# ----------------------------------------------------------------------------------------------------------------------------------- #

        # Change Entry
        total_change = StringVar(value="0.00")
        CTkLabel(master=self.edit_window, text="Change:", text_color="black",font=("Arial", 14), justify="center").place(x=10,y=310)

        self.total_change = CTkEntry(master=self.edit_window, placeholder_text="0",text_color="black",font=("Arial", 13),fg_color="transparent", border_color=self.COLOR3, height=10, textvariable=total_change)
        self.total_change.place(x=100, y=313)

# ----------------------------------------------------------------------------------------------------------------------------------- #

        # Status Entry
        status = StringVar(value=row_data[5])
        CTkLabel(master=self.edit_window, text="Status:", text_color="black",font=("Arial", 14), justify="center").place(x=265,y=253)
        self.status = CTkComboBox(master=self.edit_window, button_color=self.COLOR1, border_color=self.COLOR2, button_hover_color=self.COLOR2, width=100, dropdown_hover_color=self.COLOR3, state="readonly", values=['Serving','Pending','Dining','Completed'], variable=status)
        self.status.place(x=325,y=253)
            
# ----------------------------------------------------------------------------------------------------------------------------------- #

        # Button for edit on Edit Window
        CTkButton(master=self.edit_window, text="Edit", text_color=self.COLOR4, fg_color=self.COLOR1,font=("Arial", 15), width=90, command=lambda: self.edit_confirmation_window(order_id, self.row_data_orders, self.row)).place(x=300,y=300)

        # Close Button on Edit Window
        CTkButton(master=self.edit_window, text="Delete", text_color=self.COLOR4, fg_color="red",hover_color="#f2564b", font=("Arial", 15), width=90, command=lambda: self.delete_confirmation_window(order_id, self.row)).place(x=400,y=300)

    def create_components(self, item_name, x, y, var):
        # Check Box
        checkbox = CTkCheckBox(self.edit_window, text=item_name, width=5,
                               command=lambda: self.orders_check(slider, quantity_label, var, self.row_data_orders, item_name, self.total_bill),
                               variable=var, onvalue="on", offvalue="off")
        checkbox.place(x=x, y=y)

        quantity_label = CTkLabel(master=self.edit_window, text=f"{self.row_data_orders[item_name]}", font=("Arial", 15), justify="center")
        quantity_label.place(x=x+95, y=y-1)

        slider = CTkSlider(master=self.edit_window, from_=0, to=10,
                           command=lambda value: self.slider_event(value, quantity_label, item_name, self.row_data_orders, self.total_bill),
                           width=100, state='disabled' if var.get() == 'off' else 'normal')
        slider.place(x=x+115, y=y+5)

        return checkbox, quantity_label, slider

    # Show window when edit is clicked for confirmation to avoid misclick
    def edit_confirmation_window(self, order_id, row_data_orders, row):
        self.confirmation_window = CTkToplevel()
        self.confirmation_window.grab_set()
        self.confirmation_window.title("Confirmation")
        self.confirmation_window.after(250, lambda: self.confirmation_window.iconbitmap('assets/logo.ico'))
        self.confirmation_window.geometry(self.center_window(self.confirmation_window, 260, 100, self.confirmation_window._get_window_scaling()))
        self.confirmation_window.resizable(0, 0)

        # Text Label Window
        CTkLabel(master=self.confirmation_window, text=f"Are you sure you want to Modify\nChanges on Order #{order_id}?",font=("Arial", 15)).place(x=25,y=15)

        # Yes button, When clicked proceed to edit function in which edit the data
        CTkButton(master=self.confirmation_window, text="Yes", text_color=self.COLOR4, fg_color=self.COLOR1,font=("Arial", 15), width=90, command=lambda: self.edit(row,
                            self.table_number.get(), 
                            row_data_orders, 
                            self.total_bill.get(), 
                            self.status.get())).place(x=30,y=60)
        
        # No button, when clicked it will just close the window
        CTkButton(master=self.confirmation_window, text="No", text_color=self.COLOR4, fg_color="red", hover_color="#f2564b",font=("Arial", 15), width=90, command=lambda: self.confirmation_window.destroy()).place(x=140,y=60)

    # Edit the data in edit window
    def edit(self, row, table_number, orders, total_bill, status):
        try:
            total_bill = float(total_bill)
        except:
            messagebox.showerror("Error Message", "Invalid Input!")
            return
        order = ""
        
        # get the menu separated by comma to it in excel
        for menu, num in orders.items():
            order += (menu + ", ") * num

        # If table_number is a TakeOut or a string, just append it as a string else as a integer
        if table_number == "TakeOut":
            self.sheet[f"B{row}"] = table_number
        else:
            self.sheet[f"B{row}"] = int(table_number)
        
        # Change the value in excel
        self.sheet[f"E{row}"] = total_bill
        self.sheet[f"F{row}"] = status
        self.sheet[f"C{row}"] = order[:-2] # stop at last 2 digit to erase the ", "
        self.workbook.save(self.EXCEL_FILE)

        # destroy the window
        self.confirmation_window.destroy()
        self.edit_window.destroy()

    def delete_confirmation_window(self, order_id, row):
        self.confirmation_window = CTkToplevel()
        self.confirmation_window.grab_set()
        self.confirmation_window.title("Confirmation")
        self.confirmation_window.after(250, lambda: self.confirmation_window.iconbitmap('assets/logo.ico'))
        self.confirmation_window.geometry(self.center_window(self.confirmation_window, 260, 100, self.confirmation_window._get_window_scaling()))
        self.confirmation_window.resizable(0, 0)

        CTkLabel(master=self.confirmation_window, text=f"Are you sure you want to Delete\nOrder #{order_id} Permanently?",font=("Arial", 15)).place(x=25,y=15)

        CTkButton(master=self.confirmation_window, text="Yes", text_color=self.COLOR4, fg_color=self.COLOR1,font=("Arial", 15), width=90, command=lambda: self.delete(row)).place(x=30,y=60)

        CTkButton(master=self.confirmation_window, text="No", text_color=self.COLOR4, fg_color="red", hover_color="#f2564b",font=("Arial", 15), width=90, command=lambda: self.confirmation_window.destroy()).place(x=140,y=60)

    def delete(self, row):
        # delete the given row
        self.sheet.delete_rows(row, 1)
        self.workbook.save(self.EXCEL_FILE)

        # destroy the window
        self.confirmation_window.destroy()
        self.edit_window.destroy()

    def display_change(self, event, amount, bill):
        try:
            amount, bill = float(amount), float(bill)
            change = amount - bill
            self.total_change.delete(0, END)
            self.amount_pay.delete(0, END)
            self.total_change.insert(0, f"{change:.2f}")
            self.amount_pay.insert(0, f"{amount:.2f}")
        except:
            messagebox.showerror("Error Message", "Invalid Input!")

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