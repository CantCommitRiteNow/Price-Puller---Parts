import customtkinter as ctk
from tkinter import messagebox
import pandas as pd
import os
import webbrowser
from datetime import datetime

# Initialize
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

DB_FILE = "euro_parts_database.xlsx"
if not os.path.exists(DB_FILE):
    df = pd.DataFrame(columns=["Car", "Part Name", "Part Number", "URL", "Date Added", "Price"])
    df.to_excel(DB_FILE, index=False)
else:
    df = pd.read_excel(DB_FILE)

# Load car types from input_links.txt (unique, sorted)
CAR_LIST = []
if os.path.exists("input_links.txt"):
    car_set = set()
    with open("input_links.txt", "r") as f:
        for line in f:
            if "|" in line:
                car_set.add(line.split("|")[0].strip())
    CAR_LIST = sorted(car_set)
else:
    CAR_LIST = ["Unknown Car"]

class EuroPartsDatabase(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🧰 Euro Parts Database")
        self.geometry("1000x500")

        # Variables
        self.car_var = ctk.StringVar()
        self.url_var = ctk.StringVar()
        self.search_var = ctk.StringVar()

        # Keep full vs filtered car list
        self.full_car_list = CAR_LIST
        self.filtered_car_list = CAR_LIST.copy()

        self.create_widgets()

    def create_widgets(self):
        padding = {"padx": 10, "pady": 10}

        title = ctk.CTkLabel(self, text="Euro Parts Database", font=("Arial", 28, "bold"))
        title.pack(**padding)

        form_frame = ctk.CTkFrame(self)
        form_frame.pack(pady=10)

        # Car dropdown with live filter & auto-select
        ctk.CTkLabel(form_frame, text="Car:").grid(row=0, column=0, sticky="e", **padding)
        self.car_combo = ctk.CTkComboBox(
            form_frame,
            values=self.full_car_list,
            variable=self.car_var,
            width=300
        )
        self.car_combo.grid(row=0, column=1, **padding)
        self.car_combo.bind("<KeyRelease>", self.filter_cars)

        # URL (editable)
        ctk.CTkLabel(form_frame, text="URL:").grid(row=1, column=0, sticky="e", **padding)
        self.url_entry = ctk.CTkEntry(form_frame, textvariable=self.url_var, width=500)
        self.url_entry.grid(row=1, column=1, **padding)

        # Buttons
        add_button = ctk.CTkButton(self, text="Add Part", command=self.save_entry)
        add_button.pack(**padding)

        open_db_button = ctk.CTkButton(self, text="Open Database", command=self.open_database)
        open_db_button.pack(**padding)

        # Search box & results
        ctk.CTkLabel(self, text="Search Part Name or Number:").pack(pady=(20, 0))
        self.search_entry = ctk.CTkEntry(self, textvariable=self.search_var, width=400)
        self.search_entry.pack(**padding)
        self.search_entry.bind("<KeyRelease>", self.search_autofill)

        self.result_text = ctk.CTkTextbox(self, width=800, height=250)
        self.result_text.pack(**padding)

    def filter_cars(self, event=None):
        typed = self.car_var.get().lower()
        matches = [c for c in self.full_car_list if typed in c.lower()]
        if not matches:
            matches = ["No Match"]
        self.filtered_car_list = matches
        self.car_combo.configure(values=self.filtered_car_list)
        if len(matches) == 1 and matches[0] != "No Match":
            self.car_var.set(matches[0])

    def save_entry(self):
        car = self.car_var.get().strip()
        url = self.url_var.get().strip()

        if not car or not url:
            # keep using a label instead of popup for errors, if desired
            messagebox.showerror("Error", "Please select a car and enter a URL.")
            return

        if not url.startswith(("http://", "https://")):
            url = "https://" + url

        if "fcpeuro.com" not in url:
            messagebox.showerror("Error", "Only FCP Euro URLs are allowed.")
            return

        date_now = datetime.now().strftime("%Y-%m-%d")

        # Record with empty placeholders for scraped data
        new_row = {
            "Car": car,
            "Part Name": "",
            "Part Number": "",
            "URL": url,
            "Date Added": date_now,
            "Price": ""
        }

        global df
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(DB_FILE, index=False)

        # Instead of a popup, show confirmation in the text box
        self.result_text.insert("end", "Part Injection Successful\n")

        # Clear only the URL, keep car selection
        self.url_var.set("")

    def open_database(self):
        webbrowser.open(DB_FILE)

    def search_autofill(self, event=None):
        query = self.search_var.get().lower()
        matches = df[
            df["Part Name"].str.lower().str.contains(query) |
            df["Part Number"].str.lower().str.contains(query)
        ]

        self.result_text.delete("1.0", "end")
        if matches.empty:
            self.result_text.insert("end", "No part found.")
        else:
            for _, row in matches.iterrows():
                self.result_text.insert(
                    "end",
                    f"{row['Car']} | {row['Part Name']} | {row['Part Number']} | ${row['Price']}\n"
                )
            if len(matches) == 1:
                self.display_price_stats(matches.iloc[0]["Part Name"])

    def display_price_stats(self, part_name):
        prices = df[df["Part Name"] == part_name]["Price"].dropna().astype(float)
        if prices.empty:
            return
        low, avg, high, current = prices.min(), prices.mean(), prices.max(), prices.iloc[-1]
        first_date = pd.to_datetime(df[df["Part Name"] == part_name]["Date Added"].iloc[0])
        days = (datetime.now() - first_date).days
        change = current - low
        percent = (change / low * 100) if low else 0
        verdict = f"\nThis part has increased by ${change:.2f} ({percent:.1f}%) over {days} days."
        self.result_text.insert(
            "end",
            f"\nLow: ${low:.2f}\nAvg: ${avg:.2f}\nHigh: ${high:.2f}\nCurrent: ${current:.2f}{verdict}"
        )

if __name__ == "__main__":
    app = EuroPartsDatabase()
    app.mainloop()
