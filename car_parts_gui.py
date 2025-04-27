import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import json
import os

DATA_FILE = "car_parts.json"

# Load existing data or initialize
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return {}

def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=4)

class CarPartsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Car Parts Manager")

        self.data = load_data()

        self.car_var = tk.StringVar()
        self.part_url_var = tk.StringVar()

        # Dropdown for existing cars
        self.car_dropdown = ttk.Combobox(root, textvariable=self.car_var, state="readonly")
        self.update_dropdown()
        self.car_dropdown.grid(row=0, column=0, padx=10, pady=10)

        # Button to add a new car
        self.add_car_button = tk.Button(root, text="Add New Car", command=self.add_new_car)
        self.add_car_button.grid(row=0, column=1, padx=10)

        # Entry for part URL
        tk.Label(root, text="Part URL:").grid(row=1, column=0, padx=10, sticky="e")
        self.part_url_entry = tk.Entry(root, textvariable=self.part_url_var, width=50)
        self.part_url_entry.grid(row=1, column=1, padx=10, pady=10)

        # Submit button
        self.add_part_button = tk.Button(root, text="Add Part URL", command=self.add_part_url)
        self.add_part_button.grid(row=2, column=0, columnspan=2, pady=10)

    def update_dropdown(self):
        car_list = list(self.data.keys())
        self.car_dropdown["values"] = car_list
        if car_list:
            self.car_dropdown.current(0)

    def add_new_car(self):
        new_car = simpledialog.askstring("New Car", "Enter the car name:")
        if new_car:
            if new_car in self.data:
                messagebox.showwarning("Warning", "Car already exists!")
            else:
                self.data[new_car] = []
                self.update_dropdown()
                save_data(self.data)
                messagebox.showinfo("Success", f"{new_car} added!")

    def add_part_url(self):
        car = self.car_var.get()
        url = self.part_url_var.get().strip()

        if not car:
            messagebox.showerror("Error", "Please select a car.")
            return
        if not url:
            messagebox.showerror("Error", "Please enter a part URL.")
            return

        if url in self.data.get(car, []):
            messagebox.showwarning("Duplicate", "This URL is already added for the selected car.")
        else:
            self.data.setdefault(car, []).append(url)
            save_data(self.data)
            messagebox.showinfo("Success", "Part URL added successfully!")
            self.part_url_var.set("")  # Clear the field

if __name__ == "__main__":
    root = tk.Tk()
    app = CarPartsApp(root)
    root.mainloop()
