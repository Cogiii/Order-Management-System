from tkinter import messagebox
from resources import Resources
import matplotlib.pyplot as plt

# A function that view the charts 
class ViewChart(Resources):
    def __init__(self, choice, start_date, end_date, sheet):
        super().__init__()
        self.viewChart(choice, start_date, end_date, sheet)


    def viewChart(self, choice, starting_date, ending_date, sheet):
        start_date = starting_date.get()
        end_date = ending_date.get()
        filtered_start_date = int(start_date.replace("-", ""))
        filtered_end_date = int(end_date.replace("-", ""))

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
                    wp = {'linewidth': 1, 'edgecolor': self.COLOR1}
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
                    # display the menu and the value
                    ax.legend(wedges, [f'{label} ({value})' for label, value in zip(menu, values)],
                            title="Menu List",
                            loc="center left",
                            bbox_to_anchor=(1, 0, 0.5, 1))
                    
                    plt.setp(autotexts, size=8, weight="bold")
                    plt.title(f"Top Menu ({start_date} - {end_date})")
                    plt.show()
                else:
                    messagebox.showinfo("Error Message", "No Data Found between those dates!!!")