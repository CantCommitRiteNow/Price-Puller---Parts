🛠️ Car Parts Price Scraper
This script lets you add car part URLs to a simple text file along with the car model they belong to, and automatically scrapes the price data into a neatly organized Excel spreadsheet.

📋 How it Works
You run the script (main.py).

It asks you for:

The car model (e.g., E46 M3, 991 GT3, etc.)

The part URL (copy/paste from any website).

It stores your input into a text file (input_links.txt).

The script scrapes price and part info from the URL.

It logs the data daily into an Excel file (CarParts_Pricing.xlsx), automatically sorting everything by car model.

📂 Files Included

File	Purpose
main.py	Main script that handles scraping and Excel updates.

input_links.txt	Where your entered car parts and URLs are stored.

requirements.txt	Python packages needed to run the script.

CarParts_Pricing.xlsx	(Auto-created) Excel file storing the scraped data.

🚀 How to Run
Install dependencies
Run this once:

	bash
	Copy
	Edit
	pip install -r requirements.txt
	Start the script

	bash
	Copy
	Edit
	python main.py
	
Follow the prompts:

Enter the Car Model (like E92 M3).

Enter the Product URL.

Done! 🎉 Data will automatically appear inside the CarParts_Pricing.xlsx file.

🧹 Notes
You can keep adding more parts over time — the script will never overwrite your previous entries.

Make sure the Excel file is closed before running the script again, otherwise it can't save updates.