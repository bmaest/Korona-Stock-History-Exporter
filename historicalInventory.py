import requests
from openpyxl import Workbook
from datetime import datetime
import logging
import os
from urllib.parse import urlparse, urljoin

# Get script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Prompt user for target date
target_date_str = input("Enter target date (YYYY-MM-DD): ")
try:
    target_date = datetime.strptime(target_date_str, "%Y-%m-%d")
    target_date_str_clean = target_date.strftime('%Y%m%d')
except ValueError:
    print("❌ Invalid date format. Use YYYY-MM-DD.")
    exit()

# Prompt user for Organizational Unit
org_unit_input = input("Enter Organizational Unit name: ").strip().lower()
org_unit_safe = org_unit_input.replace(" ", "_").replace("/", "_")

# Set filenames with OU included
log_filename = os.path.join(script_dir, f"inventory_log_{org_unit_safe}_{target_date_str_clean}.log")
excel_filename = os.path.join(script_dir, f"inventory_snapshot_{org_unit_safe}_{target_date_str_clean}.xlsx")

# Setup logging
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s — %(levelname)s — %(message)s"
)

logging.info("🚀 Script started")
logging.info(f"Target date set to: {target_date.strftime('%Y-%m-%d')}")
logging.info(f"Organizational Unit filter set to: {org_unit_input}")

# API Configuration
base_url = "https://185.koronacloud.com/web/api/v3/accounts"
koronaAccountId = ""
username = ""
password = ""

# Filtering Configuration
exclude_deleted_products = True
exclude_non_tracked_products = True

# Excel Setup
workbook = Workbook()
sheet = workbook.active
sheet.title = "Inventory Snapshot"
sheet.append([
    "Product UUID", "Product Number", "Product Name",
    "Actual Stock", "Receipt Qty", "Adjustment Qty", "Sales Qty",
    "Final Stock", "Net Movement", "Inventory Timestamp"
])

# Helper: API GET with auth and logging
def get_api_data(url):
    logging.info(f"Calling API: {url}")
    try:
        response = requests.get(url, auth=(username, password))
        logging.info(f"Response Status: {response.status_code}")
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 204:
            logging.warning("No content returned.")
            return None
        else:
            logging.error(f"Error {response.status_code}: {response.text}")
            return None
    except Exception as e:
        logging.exception(f"Exception during API call: {e}")
        return None

def get_all_pages(endpoint):
    results = []
    next_url = endpoint
    parsed = urlparse(endpoint)
    domain_root = f"{parsed.scheme}://{parsed.netloc}"

    while next_url:
        data = get_api_data(next_url)
        if not data or "results" not in data:
            break
        results.extend(data["results"])
        next_link = data.get("links", {}).get("next")
        next_url = urljoin(domain_root, next_link) if next_link else None

    return results

# Step 0: Build Product Metadata
product_metadata = {}
products_url = f"{base_url}/{koronaAccountId}/products"
all_products = get_all_pages(products_url)

for product in all_products:
    product_id = product.get("id")
    if not product_id:
        continue
    product_metadata[product_id.strip()] = {
        "trackInventory": product.get("trackInventory", True),
        "name": product.get("name", "Unknown"),
        "number": product.get("number", "Unknown")
    }

# Dictionary to track product totals
product_totals = {}

def should_skip_product(product_id):
    metadata = product_metadata.get(product_id)
    if exclude_deleted_products and metadata is None:
        logging.info(f"🚫 Skipped deleted product UUID: {product_id}")
        return True
    if exclude_non_tracked_products and metadata and not metadata.get("trackInventory", True):
        logging.info(f"🚫 Skipped non-tracked product: {metadata.get('name')} ({metadata.get('number')})")
        return True
    return False

def ensure_product_entry(key, product, fallback_timestamp=None):
    if key not in product_totals:
        metadata = product_metadata.get(key, {})
        product_totals[key] = {
            "name": metadata.get("name", product.get("name", "Unknown")),
            "number": metadata.get("number", product.get("number", "Unknown")),
            "uuid": product.get("id", key),
            "actual": 0,
            "receipts": 0,
            "adjustments": 0,
            "sales": 0,
            "inventory_timestamp": fallback_timestamp
        }
        logging.warning(f"🆕 Initialized product UUID {key} with zero inventory")

# Step 1: Get Inventories
inventory_url = f"{base_url}/{koronaAccountId}/inventories"
inventories = get_all_pages(inventory_url)

for inventory in inventories:
    inventory_id = inventory.get("id")
    inventory_time_str = inventory.get("executionTime", "")
    is_active = inventory.get("active", False)
    is_booked = inventory.get("hasBookedReceipts", False)

    org_units = inventory.get("organizationalUnits", [])
    org_unit_names = [unit.get("name", "").strip().lower() for unit in org_units]

    if not is_active or not is_booked or not inventory_time_str or org_unit_input not in org_unit_names:
        logging.debug(f"⏩ Skipped inventory — Available units: {org_unit_names}")
        continue

    try:
        inventory_timestamp = datetime.strptime(inventory_time_str[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        continue

    if inventory_timestamp > target_date:
        continue

    lists_url = f"{base_url}/{koronaAccountId}/inventories/{inventory_id}/inventoryLists"
    inventory_lists = get_all_pages(lists_url)
    for inv_list in inventory_lists:
        list_id = inv_list.get("id")
        items_url = f"{base_url}/{koronaAccountId}/inventories/{inventory_id}/inventoryLists/{list_id}/items"
        items_data = get_all_pages(items_url)
        for item in items_data:
            product = item.get("product", {})
            stock = item.get("stock", {})
            product_id = product.get("id")
            actual_stock = stock.get("actual")

            if not product_id or actual_stock is None:
                continue

            key = product_id.strip()
            if should_skip_product(key):
                continue

            existing = product_totals.get(key)
            if not existing or inventory_timestamp > existing["inventory_timestamp"]:
                ensure_product_entry(key, product, inventory_timestamp)
                product_totals[key]["actual"] = actual_stock
                product_totals[key]["inventory_timestamp"] = inventory_timestamp
                logging.info(f"📦 Inventory recorded for UUID {key} at {inventory_timestamp.isoformat()}")

# Step 2: Get Stock Receipts
receipts_url = f"{base_url}/{koronaAccountId}/stockReceipts"
receipts = get_all_pages(receipts_url)

for receipt in receipts:
    if receipt.get("status") != "BOOKED":
        continue

    booking_time_str = receipt.get("bookingTime", "")
    org_unit = receipt.get("organizationalUnit", {}).get("name", "").strip().lower()
    if not booking_time_str or org_unit != org_unit_input:
        logging.debug(f"⏩ Skipped receipt for unit '{org_unit}'")
        continue

    try:
        booking_timestamp = datetime.strptime(booking_time_str[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        continue

    if booking_timestamp > target_date:
        continue

    receipt_id = receipt.get("id")
    items_url = f"{base_url}/{koronaAccountId}/stockReceipts/{receipt_id}/items"
    items_data = get_all_pages(items_url)
    for item in items_data:
        product = item.get("product", {})
        product_id = product.get("id")
        amount_received = item.get("amount", {}).get("received")

        if not product_id or amount_received is None:
            continue

        key = product_id.strip()
        if should_skip_product(key):
            continue

        ensure_product_entry(key, product)
        inventory_timestamp = product_totals[key]["inventory_timestamp"]

        if inventory_timestamp is None or inventory_timestamp < booking_timestamp <= target_date:
            product_totals[key]["receipts"] += amount_received
            logging.info(
                f"📥 Receipt added — Time: {booking_timestamp.isoformat()}, "
                f"Product: {product.get('name', 'Unknown')} ({product.get('number')}), "
                f"Amount: +{amount_received}"
            )

# Step 3: Get Stock Adjustments
adjustments_url = f"{base_url}/{koronaAccountId}/stockAdjustments"
adjustments = get_all_pages(adjustments_url)

for adjustment in adjustments:
    if adjustment.get("status") != "BOOKED":
        continue

    booking_time_str = adjustment.get("bookingTime", "")
    org_unit = adjustment.get("warehouse", {}).get("name", "").strip().lower()

    if not booking_time_str or org_unit != org_unit_input:
        logging.debug(f"⏩ Skipped adjustment for unit '{org_unit}'")
        continue

    try:
        booking_timestamp = datetime.strptime(booking_time_str[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        continue

    if booking_timestamp > target_date:
        continue

    adjustment_id = adjustment.get("id")
    items_url = f"{base_url}/{koronaAccountId}/stockAdjustments/{adjustment_id}/items"
    items_data = get_all_pages(items_url)
    for item in items_data:
        product = item.get("product", {})
        product_id = product.get("id")
        amount = item.get("amount")

        if not product_id or amount is None:
            continue

        key = product_id.strip()
        if should_skip_product(key):
            continue

        ensure_product_entry(key, product)
        inventory_timestamp = product_totals[key]["inventory_timestamp"]

        if inventory_timestamp is None or inventory_timestamp < booking_timestamp <= target_date:
            product_totals[key]["adjustments"] += amount
            logging.info(
                f"🛠️ Adjustment added — Time: {booking_timestamp.isoformat()}, "
                f"Product: {product.get('name', 'Unknown')} ({product.get('number')}), "
                f"Amount: {amount:+}"
            )

# Step 4: Get Sales Receipts
sales_url = f"{base_url}/{koronaAccountId}/receipts"
sales_receipts = get_all_pages(sales_url)

for receipt in sales_receipts:
    if receipt.get("cancelled") or receipt.get("voided"):
        continue

    org_unit = receipt.get("organizationalUnit", {}).get("name", "").strip().lower()
    if org_unit != org_unit_input:
        logging.debug(f"⏩ Skipped sales receipt for unit '{org_unit}'")
        continue

    booking_time_str = receipt.get("bookingTime", "")
    if not booking_time_str:
        continue

    try:
        booking_timestamp = datetime.strptime(booking_time_str[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        continue

    if booking_timestamp > target_date:
        continue

    receipt_id = receipt.get("id")
    detail_url = f"{base_url}/{koronaAccountId}/receipts/{receipt_id}"
    receipt_detail = get_api_data(detail_url)
    if not receipt_detail or not receipt_detail.get("items"):
        continue

    for item in receipt_detail["items"]:
        if item.get("type") != "PRODUCT":
            continue

        product = item.get("product", {})
        product_id = product.get("id")
        quantity_sold = item.get("quantity")

        if not product_id or quantity_sold is None:
            continue

        key = product_id.strip()
        if should_skip_product(key):
            continue

        ensure_product_entry(key, product)
        inventory_timestamp = product_totals[key]["inventory_timestamp"]

        if inventory_timestamp is None or inventory_timestamp < booking_timestamp <= target_date:
            product_totals[key]["sales"] -= quantity_sold
            logging.info(
                f"🧾 Sale recorded — Time: {booking_timestamp.isoformat()}, "
                f"Product: {product.get('name', 'Unknown')} ({product.get('number')}), "
                f"Amount: -{quantity_sold}"
            )

# Step 5: Write to Excel
excluded_deleted_count = 0
excluded_non_tracked_count = 0
included_count = 0

for product in product_totals.values():
    key = product["uuid"]
    if should_skip_product(key):
        metadata = product_metadata.get(key, {})
        if metadata is None:
            excluded_deleted_count += 1
        elif not metadata.get("trackInventory", True):
            excluded_non_tracked_count += 1
        continue

    final_stock = product["actual"] + product["receipts"] + product["adjustments"] + product["sales"]
    net_movement = product["receipts"] + product["adjustments"] + product["sales"]

    sheet.append([
        product["uuid"],
        product["number"],
        product["name"],
        product["actual"],
        product["receipts"],
        product["adjustments"],
        product["sales"],
        final_stock,
        net_movement,
        product["inventory_timestamp"].strftime("%Y-%m-%d %H:%M:%S") if product["inventory_timestamp"] else "",
    ])

    included_count += 1

# Save Excel
try:
    workbook.save(excel_filename)
    logging.info(f"✅ Inventory snapshot saved to {excel_filename}")
    logging.info(f"📦 Included products: {included_count}")
    logging.info(f"🗑️ Excluded deleted products: {excluded_deleted_count}")
    logging.info(f"🚫 Excluded non-tracked products: {excluded_non_tracked_count}")
except Exception as e:
    logging.exception(f"❌ Failed to save Excel file: {e}")

logging.info("🏁 Script completed")
print(f"\n✅ Inventory snapshot saved to: {excel_filename}")
print(f"📄 Log file saved to: {log_filename}")
