import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import platform

DATABASE_FILE = 'euro_parts_database.xlsx'

class EuroPartsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Euro Parts Database")
        self.geometry("1000x750")
        
        default_font = ("Segoe UI", 11)
        self.option_add('*Font', default_font)

        self.configure_theme()

        self.parts_data = self.load_database()

        self.create_widgets()

    def configure_theme(self):
        try:
            if platform.system() == "Windows":
                self.tk.call("source", "azure.tcl")
                self.tk.call("set_theme", "dark")
            elif platform.system() == "Darwin":
                self.tk.call("source", "azure.tcl")
                self.tk.call("set_theme", "dark")
        except:
            pass  # If Azure theme isn't installed, just ignore

    def load_database(self):
        if os.path.exists(DATABASE_FILE):
            return pd.read_excel(DATABASE_FILE)
        else:
            return pd.DataFrame(columns=['Car', 'Part Name', 'Part Number', 'URL', 'Price', 'Date Added'])

    def create_widgets(self):
        frame_top = ttk.Frame(self)
        frame_top.pack(pady=15)

        # Car Dropdown
        ttk.Label(frame_top, text="Car:").grid(row=0, column=0, sticky='e')
        self.car_var = tk.StringVar()
        self.car_dropdown = ttk.Combobox(frame_top, textvariable=self.car_var, values=self.get_unique_cars(), width=35)
        self.car_dropdown.grid(row=0, column=1, padx=8)

        # Part Name Entry
        ttk.Label(frame_top, text="Part Name:").grid(row=1, column=0, sticky='e')
        self.part_var = tk.StringVar()
        self.part_entry = ttk.Entry(frame_top, textvariable=self.part_var, width=35)
        self.part_entry.grid(row=1, column=1, padx=8)
        self.part_var.trace("w", self.autofill_part_name)

        # Part Number Entry
        ttk.Label(frame_top, text="Part Number:").grid(row=2, column=0, sticky='e')
        self.partnum_var = tk.StringVar()
        self.partnum_entry = ttk.Entry(frame_top, textvariable=self.partnum_var, width=35)
        self.partnum_entry.grid(row=2, column=1, padx=8)
        self.partnum_var.trace("w", self.autofill_part_number)

        # URL Entry
        ttk.Label(frame_top, text="URL:").grid(row=3, column=0, sticky='e')
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(frame_top, textvariable=self.url_var, width=35)
        self.url_entry.grid(row=3, column=1, padx=8)

        # Search Entry
        ttk.Label(frame_top, text="Search Part:").grid(row=4, column=0, sticky='e')
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(frame_top, textvariable=self.search_var, width=35)
        self.search_entry.grid(row=4, column=1, padx=8)
        self.search_var.trace("w", self.search_part)

        # Buttons
        frame_buttons = ttk.Frame(self)
        frame_buttons.pack(pady=20)

        ttk.Button(frame_buttons, text="➕ Add Part", command=self.add_part).grid(row=0, column=0, padx=12)
        ttk.Button(frame_buttons, text="📂 Open Database", command=self.open_database).grid(row=0, column=1, padx=12)

        # Output Text
        default_font = ("Segoe UI", 11)
        self.option_add('*Font', default_font)
        self.output_text.pack(pady=20)

    def get_unique_cars(self):
        return sorted(self.parts_data['Car'].dropna().unique().tolist())

    def autofill_part_name(self, *args):
        text = self.part_var.get().lower()
        matches = self.parts_data[self.parts_data['Part Name'].str.lower().str.contains(text, na=False)]
        if not matches.empty:
            first_match = matches.iloc[0]
            self.car_var.set(first_match['Car'])
            self.partnum_var.set(first_match.get('Part Number', ''))

    def autofill_part_number(self, *args):
        text = self.partnum_var.get().lower()
        matches = self.parts_data[self.parts_data['Part Number'].astype(str).str.lower().str.contains(text, na=False)]
        if not matches.empty:
            first_match = matches.iloc[0]
            self.car_var.set(first_match['Car'])
            self.part_var.set(first_match.get('Part Name', ''))

    def search_part(self, *args):
        query = self.search_var.get().lower()
        matches = self.parts_data[self.parts_data.apply(lambda row: query in str(row['Part Name']).lower() or query in str(row['Part Number']).lower(), axis=1)]

        self.output_text.delete('1.0', tk.END)
        if not matches.empty:
            for _, row in matches.iterrows():
                self.output_text.insert(tk.END, f"Car: {row['Car']} | Part: {row['Part Name']} | Part #: {row['Part Number']} | URL: {row['URL']}\n")
        else:
            self.output_text.insert(tk.END, "No matching part found.\n")

    def add_part(self):
        car = self.car_var.get()
        part_name = self.part_var.get()
        part_number = self.partnum_var.get()
        url = self.url_var.get()

        if not url.startswith("http"):
            url = "https://" + url

        if "fcpeuro.com" not in url:
            messagebox.showerror("Invalid URL", "Only URLs from FCP Euro are allowed.")
            return

        if not all([car, part_name, part_number, url]):
            messagebox.showerror("Missing Info", "Please fill out all fields.")
            return

        new_entry = pd.DataFrame([{
            'Car': car,
            'Part Name': part_name,
            'Part Number': part_number,
            'URL': url,
            'Price': None,
            'Date Added': pd.Timestamp.now()
        }])

        self.parts_data = pd.concat([self.parts_data, new_entry], ignore_index=True)
        self.parts_data.to_excel(DATABASE_FILE, index=False)

        messagebox.showinfo("Success", "Part added successfully.")
        self.clear_inputs()

    def clear_inputs(self):
        self.car_var.set('')
        self.part_var.set('')
        self.partnum_var.set('')
        self.url_var.set('')
        self.search_var.set('')
        self.output_text.delete('1.0', tk.END)

    def open_database(self):
        os.system(f'start excel "{DATABASE_FILE}"' if platform.system() == "Windows" else f'open "{DATABASE_FILE}"')

if __name__ == "__main__":
    app = EuroPartsApp()
    app.mainloop()