from customtkinter import *

app = CTk()
app.geometry("900x700")
app.resizable(0,0)
app.title("Order System")
set_appearance_mode("light")

sidebar_frame = CTkFrame(master=app, fg_color="#2A8C55",  width=176, height=650, corner_radius=0)
sidebar_frame.pack_propagate(0)
sidebar_frame.pack(fill="y", anchor="w", side="left")

app.mainloop()