# price_scraper/scraper.py

import requests
import json
from lxml import html
import logging

logging.basicConfig(
    filename='price_puller.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

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

        name_from_site = data.get('name', product_name) or product_name
        offers = data.get('offers', {})
        if isinstance(offers, list):
            offers = offers[0]

        try:
            price = float(offers.get('price', "0"))
        except (ValueError, TypeError):
            price = None

        sku = offers.get('sku', "N/A")

        return {
            "Name": name_from_site,
            "SKU": sku,
            "Price": price,
            "URL": url
        }

    except Exception as e:
        logging.error(f"🚨 Error fetching data for {product_name}: {e}")
        return None
