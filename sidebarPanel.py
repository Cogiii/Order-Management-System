from customtkinter import *
from PIL import Image
from resources import Resources

class SideBarPanel(Resources):
    def __init__(self, master, dashboard, new_order):
        super().__init__()
        self.show_dashboard = dashboard
        self.show_new_order = new_order

        self.sidebar_frame = CTkFrame(master=master, fg_color=self.COLOR1,  width=176, corner_radius=0)
        self.sidebar_frame.pack_propagate(0)
        self.sidebar_frame.pack(fill="y", anchor="w", side="left")

        # Display the image Logo
        logo_img_data = Image.open("assets/logo.png")
        logo_img = CTkImage(dark_image=logo_img_data, light_image=logo_img_data, size=(160, 160))
        CTkLabel(master=self.sidebar_frame, text="", image=logo_img).pack(pady=(38, 0), anchor="center")

        # Navigation Bar Buttons
        self.dashboard_nav = self.sidebarNavigation(self.sidebar_frame, "assets/analytics_icon.png", "Dashboard", 60, self.clicked_dashboard)
        self.new_order_nav = self.sidebarNavigation(self.sidebar_frame, "assets/cart_icon.png", "New Order", 16, self.clicked_new_orders)
        self.admin_nav = self.sidebarNavigation(self.sidebar_frame, "assets/person_icon.png", "Admin", 300)

    def sidebarNavigation(self, master, logo, text, pady, call_nav=None):
        # Dashboard Nav Bar
        image_data = Image.open(logo)
        image = CTkImage(dark_image=image_data, light_image=image_data)

        sidebar_nav = CTkButton(master=master, image=image, text=text, fg_color="transparent", font=("Arial Bold", 14), hover_color=self.COLOR2, anchor="w", command=call_nav)
        sidebar_nav.pack(anchor="center", ipady=5, pady=(pady, 0))

        return sidebar_nav
    
    def clicked_dashboard(self):
        self.dashboard_nav.configure(fg_color=self.COLOR2, state='disabled')
        self.new_order_nav.configure(fg_color="transparent", state='normal')
        self.show_dashboard()

    def clicked_new_orders(self):
        self.new_order_nav.configure(fg_color=self.COLOR2, state='disabled')
        self.dashboard_nav.configure(fg_color="transparent", state='normal')
        self.show_new_order()

        
