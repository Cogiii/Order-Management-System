"""
POS: Order Management System
Developed by Laurence Kharl Devera
Computer Science Student | 1st Year College 2023-2024 | Aspiring Developer
From Mapúa Malayan Colleges Mindanao
"""
from customtkinter import *
from sidebarPanel import SideBarPanel
from resources import Resources
from dashboard import Dashboard
from newOrder import NewOrder

class OMSApp(Resources):
    def __init__(self):
        self.app = CTk()
        self.app.geometry(self.center_window(self.app, 1300, 700, self.app._get_window_scaling()))
        self.app.resizable(0,0)
        self.app.title("Order Management System")
        set_appearance_mode("light")
        self.app.iconbitmap("assets/logo.ico")

        self.dashboard = Dashboard()
        self.new_order = NewOrder()

        self.show_sidebar()
        self.show_dashboard()
        
    def show_sidebar(self):
        SideBarPanel(self.app, lambda: self.show_dashboard(), lambda:self.show_new_order())

    def show_dashboard(self):
        self.dashboard.show_dashboard(self.app)
        self.new_order.hide_new_order()

    def show_new_order(self):
        self.new_order.show_new_order(self.app)
        self.dashboard.hide_dashboard()

    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    app = OMSApp()
    app.run()
