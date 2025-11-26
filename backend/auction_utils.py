import os
import csv
import re
import json
from datetime import datetime
from google.cloud import firestore
from google.oauth2 import service_account

# Path to your Firebase service account key
SERVICE_ACCOUNT_PATH = "/Users/tinman/Projects/FreeForeclosureList/backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json"
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH)
db = firestore.Client(credentials=credentials)

def process_quicksearch_to_auctions(csv_path, county, sale_type):
    """
    Process a QuickSearch CSV and upsert sales into Firestore for the given county and sale_type.
    On success, update the auction_parameters table to record the update time for this county/sale_type.
    Accepts any filename; uses today's date if no date is found in the filename.
    """
    import datetime
    base = os.path.basename(csv_path)
    m = re.match(r"[A-Za-z]+[-_](\d{8})[_-]?(Foreclosure|TaxDeed)?[_-]?QuickSearch\.csv", base, re.IGNORECASE)
    date = m.group(1) if m else datetime.datetime.now().strftime('%Y%m%d')
    total = 0
    written = 0
    skipped = 0
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        reader.fieldnames = [fn.strip() for fn in reader.fieldnames]
        for row in reader:
            row = {k.strip(): v for k, v in row.items()}
            total += 1
            casenumber = row.get("Case Number") or row.get("CaseNumber") or row.get("CaseNo")
            if not casenumber:
                skipped += 1
                continue
            written += 1
            doc_name = f"{county}_{date}"
            if sale_type and sale_type.lower() != 'unknown':
                doc_name += f"_{sale_type.lower()}"
            doc_ref = db.collection("Auctions").document(doc_name).collection("Sales").document(str(casenumber))
            existing = doc_ref.get()
            if existing.exists:
                old = existing.to_dict()
                update = {k: v for k, v in row.items() if k not in ("Case Number", "CaseNumber", "CaseNo") and old.get(k) != v}
                if update:
                    doc_ref.update(update)
            else:
                doc_ref.set(row)
    param_doc = db.collection("auction_parameters").document(f"{county}_{sale_type}")
    param_doc.set({"last_update": datetime.datetime.utcnow()}, merge=True)
    return {"total": total, "written": written, "skipped": skipped, "county": county, "sale_type": sale_type, "date": date}

def parse_date(date_str):
    try:
        return datetime.strptime(date_str.strip(), '%m/%d/%Y')
    except Exception:
        return None

def get_next_sale_for_county(county_name, json_path='backend/Legacy/foreclosureSales_clean.json'):
    """
    Returns (sale_type, sale_date, row) for the next sale for the given county.
    sale_type is 'Foreclosure' or 'Tax Deed'.
    sale_date is MM/DD/YYYY string or None.
    row is the full dict from the JSON.
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for row in data:
        if row.get('County', '').strip().lower() == county_name.strip().lower():
            fc_flag = str(row.get('UiPath', '')).strip().lower() == 'x'
            td_flag = str(row.get('UiPathTD', '')).strip().lower() == 'x'
            fc_date = parse_date(row.get('Foreclosure', '')) if fc_flag else None
            td_date = parse_date(row.get('Tax Deed', '')) if td_flag else None
            if fc_date and td_date:
                if fc_date <= td_date:
                    return 'Foreclosure', fc_date.strftime('%m/%d/%Y'), row
                else:
                    return 'Tax Deed', td_date.strftime('%m/%d/%Y'), row
            elif fc_date:
                return 'Foreclosure', fc_date.strftime('%m/%d/%Y'), row
            elif td_date:
                return 'Tax Deed', td_date.strftime('%m/%d/%Y'), row
            else:
                return None, None, row
    return None, None, None

def get_next_sale_for_county_if_stale(county_name, sale_type, min_hours=0, json_path='backend/Legacy/foreclosureSales_clean.json'):
    """
    Returns (sale_type, sale_date, row) for the next sale for the given county ONLY IF the last update for that county/sale_type
    in Firestore was more than min_hours ago. Otherwise returns (None, None, None).
    """
    from datetime import timezone, timedelta, datetime
    county_doc = db.collection("auction_parameters").document(county_name.lower()).get()
    if county_doc.exists:
        doc_dict = county_doc.to_dict()
        sale_type_data = doc_dict.get(sale_type.lower(), {})
        last_update = sale_type_data.get("last_update")
        if last_update:
            if hasattr(last_update, 'replace'):
                last_update_dt = last_update.replace(tzinfo=timezone.utc)
            else:
                last_update_dt = datetime.fromisoformat(str(last_update))
            now = datetime.now(timezone.utc)
            hours_since = (now - last_update_dt).total_seconds() / 3600.0
            if hours_since < min_hours:
                return None, None, None
    return get_next_sale_for_county(county_name, json_path)

def upload_report_and_mark_updated(report_path, county, sale_type, dest_dir='dist/reports'):
    """
    Copies the report HTML to the deploy directory and updates Firestore auction_parameters last_update.
    After this, run `firebase deploy --only hosting` to push to Firebase Hosting.
    """
    import shutil
    import datetime
    import os
    dest_path = os.path.join(dest_dir, os.path.basename(report_path))
    if os.path.abspath(report_path) != os.path.abspath(dest_path):
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)
        shutil.copy2(report_path, dest_path)
    # Always update Firestore last_update as a nested field under the sale_type in the county document
    county_doc = db.collection("auction_parameters").document(county.lower())
    update_data = {f"{sale_type.lower()}.last_update": datetime.datetime.utcnow()}
    county_doc.set(update_data, merge=True)
    # NOTE: You must still run `firebase deploy --only hosting` to upload to Firebase Hosting
    return dest_path