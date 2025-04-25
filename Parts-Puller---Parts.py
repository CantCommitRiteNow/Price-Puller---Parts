import requests
from lxml import html
import json
import logging
from datetime import datetime
import pytz
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# Setup logging
logging.basicConfig(
    filename='price_puller.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Set timezone to EST
est = pytz.timezone('US/Eastern')

def get_product_info(url, product_name):
    headers = {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "user-agent": "Mozilla/5.0"
    }

    try:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            logging.warning(f"❌ Failed to fetch page for {product_name} | Status: {response.status_code}")
            return None

        tree = html.fromstring(response.content)
        script_content = tree.xpath('//script[@type="application/ld+json"]/text()')
        if not script_content:
            logging.warning(f"⚠️ No JSON-LD found for {product_name}")
            return None

        data = json.loads(script_content[0])
        offers = data.get('offers', {})
        if isinstance(offers, list):
            offers = offers[0]

        try:
            price = float(offers.get('price', "0"))
        except (ValueError, TypeError):
            price = None

        sku = offers.get('sku', "N/A")

        return {
            "Name": product_name,
            "SKU": sku,
            "Price": price,
            "URL": url
        }

    except Exception as e:
        logging.error(f"🚨 Error fetching data for {product_name}: {e}")
        return None

def merge_price_label_row(ws, start_col, product_count):
    from_col = get_column_letter(start_col)
    to_col = get_column_letter(start_col + product_count - 1)
    ws.merge_cells(f'{from_col}3:{to_col}3')
    ws[f'{from_col}3'] = "Price (USD)"
    ws[f'{from_col}3'].font = Font(bold=True)
    ws[f'{from_col}3'].alignment = Alignment(horizontal='center')

def write_to_excel(section_name, product_results, today_str, file_path='CarParts_Pricing.xlsx'):
    base_date = datetime(2025, 4, 24)
    today_date = datetime.strptime(today_str, "%m/%d/%Y")
    days_since = (today_date - base_date).days
    row_index = 4 + days_since

    if os.path.exists(file_path):
        wb = load_workbook(file_path)
    else:
        wb = Workbook()
        # Remove default Sheet if it exists
        default_sheet = wb.active
        if default_sheet.title == "Sheet":
            wb.remove(default_sheet)

    if section_name in wb.sheetnames:
        ws = wb[section_name]
    else:
        ws = wb.create_sheet(section_name)
        ws['A1'] = 'Name'
        ws['A2'] = 'Part # / SKU'
        ws['A3'] = 'Date'
        for r in range(1, 4):
            ws.cell(row=r, column=1).font = Font(bold=True)
            ws.cell(row=r, column=1).alignment = Alignment(horizontal='right')

    start_col = 2
    for idx, product in enumerate(product_results):
        name = product["Name"]
        sku = product["SKU"]
        price = product["Price"]
        url = product["URL"]
        col = start_col + idx

        # Name with hyperlink
        name_cell = ws.cell(row=1, column=col)
        name_cell.value = name
        if url:
            name_cell.hyperlink = url
            name_cell.style = "Hyperlink"

        # SKU
        ws.cell(row=2, column=col).value = sku

        # Price with color-coded change detection
        price_cell = ws.cell(row=row_index, column=col)
        price_cell.value = price
        price_cell.number_format = '"$"#,##0.00'

        previous_price_cell = ws.cell(row=row_index - 1, column=col)
        previous_price = previous_price_cell.value

        if previous_price is not None and isinstance(previous_price, (int, float)) and isinstance(price, (int, float)):
            if price > previous_price:
                color = "FF0000"  # red
            elif price < previous_price:
                color = "00B050"  # green
            else:
                color = "000000"  # black
            price_cell.font = Font(color=color)
        else:
            price_cell.font = Font(color="000000")  # default black

    # Merge and center Price (USD)
    merge_price_label_row(ws, start_col, len(product_results))

    # Insert date in A column
    ws.cell(row=row_index, column=1).value = today_date
    ws.cell(row=row_index, column=1).number_format = 'mm/dd/yyyy'

    # Auto-fit columns based on Row 1 content
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value:
            width = len(str(cell_value)) + 2
            col_letter = get_column_letter(col)
            ws.column_dimensions[col_letter].width = width

    try:
        wb.save(file_path)
        logging.info(f"✅ Excel updated for section: {section_name}")
    except PermissionError:
        print("❌ Excel File open — please close the file and try again.")
        logging.error("❌ Excel File open — save failed.")

def main():
    today_str = datetime.now(est).strftime("%m/%d/%Y")

    car_parts = {
        "E46 M3": [
            ("https://www.fcpeuro.com/products/bmw-clutch-flywheel-z4-m3-dmf050", "BMW Z4 M3 Clutch Flywheel"),
            ("https://www.fcpeuro.com/products/bmw-clutch-kit-m3-e46-luk-21212282667", "BMW E46 M3 Clutch Kit"),
            ("https://www.fcpeuro.com/products/bmw-clutch-slave-cylinder-oem-21526785966", "BMW Clutch Slave Cylinder OEM"),
            ("https://www.fcpeuro.com/products/bmw-clutch-pilot-bearing-11211720310", "BMW Clutch Pilot Bearing"),
            ("https://www.fcpeuro.com/products/bmw-differential-cover-kit-33112282482kt", "BMW Differential Cover Kit"),
            ("https://www.fcpeuro.com/products/bmw-oil-change-kit-e46-m3-z3m-z4m-11427833769kt2", "BMW Oil Change Kit"),
            ("https://www.fcpeuro.com/products/bmw-cabin-air-filter-64319257504-1", "BMW Cabin Air Filter"),
            ("https://www.fcpeuro.com/products/bmw-wiper-blade-set-bosch-61610037009", "BMW Wiper Blade Set"),
            ("https://www.fcpeuro.com/products/jectron-fuel-injection-cleaner-liqui-moly-lm2007", "Fuel Injection Cleaner"),
            ("https://www.fcpeuro.com/products/bmw-emblem-rear-trunk-lid-51148219237", "Rear Trunk Lid Emblem"),
            ("https://www.fcpeuro.com/products/fuel-system-cleaner-500ml-can-liqui-moly-lm2030", "Fuel System Cleaner"),
        ],
        "991 GT3": [
            ("https://www.fcpeuro.com/products/porsche-center-lock-wheel-nut-kit-genuine-porsche-99136108190kt4"),
            ("https://www.fcpeuro.com/products/porsche-wheel-bearing-kit-aftermarket-95834190100"),
            ("https://www.fcpeuro.com/products/bosch-windshield-wiper-set-bmw-porsche-land-rover-bosch-3397007697"),
            ("https://www.fcpeuro.com/products/porsche-engine-oil-change-kit-5w-40-liqui-moly-991gt3oilkt7"),
            ("https://www.fcpeuro.com/products/volvo-mercedes-benz-serpentine-belt-slk230-960-s90-v90-6pk1755"),
            ("https://www.fcpeuro.com/products/porsche-air-filter-hengst-e1722l"),
            ("https://www.fcpeuro.com/products/porsche-cabin-air-filter-hengst-e3940lb"),
            ("https://www.fcpeuro.com/products/porsche-cabin-air-filter-hengst-e3945lb"),
            ("https://www.fcpeuro.com/products/porsche-brake-kit-sebro-991gt3brkt8"),
            ("https://www.fcpeuro.com/products/porsche-center-lock-wheel-socket-cta-5016"),
            ("https://www.fcpeuro.com/products/porsche-center-lock-wheel-center-cap-removal-tool-genuine-porsche-99136107900"),
            ("https://www.fcpeuro.com/products/porsche-dual-clutch-transmission-service-kit-liqui-moly-9g132102500kt1")
        ],
        "9th Gen Accord": [
            ("https://www.fcpeuro.com/products/headlight-light-bulb-upgrade-philips-crystalvision-h11cvps2", "Headlight Upgrade Bulb")
        ]
    }

    for section, items in car_parts.items():
        print(f"\n📦 Processing section: {section}")
        product_data = []

        for item in items:
            if isinstance(item, tuple):
                url, name = item
            else:
                url = item
                name = "Unnamed Product"

            data = get_product_info(url, name)
            if data:
                product_data.append(data)

        if product_data:
            write_to_excel(section, product_data, today_str)

if __name__ == '__main__':
    main()
