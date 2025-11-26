print("=== auction_utils.py is running ===")

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
    import datetime
    import os
    import re
    import csv

    # --- DEBUG: Print Firestore project and test write ---
    project_id = db.project
    print(f"[DEBUG] Firestore project_id: {project_id}")
    print(f"[DEBUG] Writing to collection: 'sales'")
    # Minimal test write
    try:
        test_doc = {'test': True, 'timestamp': datetime.datetime.now().isoformat()}
        result = db.collection('sales').add(test_doc)
        print(f"[DEBUG] Test write to 'sales' collection succeeded. Doc id: {result.id if hasattr(result, 'id') else result}")
    except Exception as e:
        print(f"[DEBUG] Test write to 'sales' collection failed: {e}")
    # --- END DEBUG ---

    base = os.path.basename(csv_path)
    m = re.match(r"[A-Za-z]+[-_](\d{8})[_-]?(Foreclosure|TaxDeed)?[_-]?QuickSearch\.csv", base, re.IGNORECASE)
    date = m.group(1) if m else datetime.datetime.now().strftime('%Y%m%d')
    total = 0
    written = 0
    skipped = 0
    print(f"Processing CSV: {csv_path} for County: {county}, Sale Type: {sale_type}")
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
            # Save each sale as a document in the top-level 'sales' collection
            sale_doc = dict(row)
            sale_doc['County'] = county
            sale_doc['SaleType'] = sale_type
            sale_doc['UploadDate'] = date
            print(f"Uploading case {casenumber} to 'sales' collection with County={county}, SaleType={sale_type}, UploadDate={date}")
            doc_ref = db.collection("sales").document(str(casenumber))
            existing = doc_ref.get()
            try:
                if existing.exists:
                    old = existing.to_dict()
                    update = {k: v for k, v in sale_doc.items() if k not in ("Case Number", "CaseNumber", "CaseNo") and old.get(k) != v}
                    if update:
                        doc_ref.update(update)
                else:
                    doc_ref.set(sale_doc)
            except Exception as e:
                print(f"Error writing case {casenumber} to Firestore: {e}")
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
    Returns (sale_type, sale_date, row, sale_list_url) for the next sale for the given county ONLY IF the last update for that county/sale_type
    in Firestore was more than min_hours ago. Otherwise returns (None, None, None, None).
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
                return None, None, None, None
    # Always return 4 values
    result = get_next_sale_for_county(county_name, json_path)
    if result is None:
        return None, None, None, None
    if len(result) == 4:
        return result
    elif len(result) == 3:
        return result[0], result[1], result[2], None
    else:
        return None, None, None, None

def upload_report_and_mark_updated(report_path, county, sale_type, dest_dir='dist/reports'):
    """
    Runs the build step, copies the report HTML to the Firebase Hosting dist directory,
    updates Firestore auction_parameters last_update, and deploys to Firebase Hosting.
    """
    import shutil
    import datetime
    import os
    import subprocess

    # 1. Run the build step
    try:
        build_result = subprocess.run(
            ["npm", "run", "build"],
            capture_output=True, text=True, check=True
        )
        build_output = build_result.stdout
        build_error = build_result.stderr
    except Exception as e:
        build_output = ""
        build_error = str(e)

    # 2. Always copy to Firebase Hosting dist directory
    firebase_dist_dir = os.path.join(os.path.dirname(__file__), '..', 'dist', 'reports')
    if not os.path.exists(firebase_dist_dir):
        os.makedirs(firebase_dist_dir)
    dest_path = os.path.join(firebase_dist_dir, os.path.basename(report_path))
    if os.path.abspath(report_path) != os.path.abspath(dest_path):
        shutil.copy2(report_path, dest_path)

    # 3. Always update Firestore last_update as a nested field under the sale_type in the county document
    county_doc = db.collection("auction_parameters").document(county.lower())
    update_data = {f"{sale_type.lower()}.last_update": datetime.datetime.utcnow()}
    county_doc.set(update_data, merge=True)

    # 4. Deploy to Firebase Hosting
    try:
        result = subprocess.run(
            ["firebase", "deploy", "--only", "hosting"],
            capture_output=True, text=True, check=True
        )
        deploy_output = result.stdout
        deploy_error = result.stderr
    except Exception as e:
        deploy_output = ""
        deploy_error = str(e)

    return {
        "build_output": build_output,
        "build_error": build_error,
        "copied_to": dest_path,
        "firestore_updated": True,
        "deploy_output": deploy_output,
        "deploy_error": deploy_error
    }


def filter_sales(sales, county, sales_type):
    filtered = []
    for row in sales:
        row_county = (row.get('County') or '').strip().lower()
        row_type = (row.get('SaleType') or row.get('Sales Type') or '').strip().lower()
        if row_county == county.lower() and row_type == sales_type.lower():
            filtered.append(row)
    return filtered

def generate_html_report_from_firestore(county, sales_type, output_path):
    # Uses the global db object (already created at the top of your file)
    sales_ref = db.collection('sales')
    docs = sales_ref.stream()
    sales = [doc.to_dict() for doc in docs]
    filtered = filter_sales(sales, county, sales_type)
    generate_html_report_from_sales(filtered, output_path, county, sales_type)

def generate_html_report_from_sales(sales, output_path, county, sales_type):
    import os
    from datetime import datetime
    # Safety check: ensure sortable-table.js exists
    js_path = os.path.join(os.path.dirname(__file__), '..', 'public', 'sortable-table.js')
    if not os.path.isfile(js_path):
        print(f"ERROR: Required JS file not found: {js_path}\nReport generation aborted to prevent loss of interactivity.")
        return

    # Load PA template for the county
    import json
    pa_template = None
    pa_template_path = os.path.join(os.path.dirname(__file__), 'Legacy', 'foreclosureSales_clean.json')
    try:
        with open(pa_template_path, 'r', encoding='utf-8') as f:
            pa_data = json.load(f)
            for entry in pa_data:
                if entry.get('County', '').strip().lower() == county.strip().lower():
                    pa_template = entry.get('PA template', '').strip()
                    break
    except Exception as e:
        pa_template = None

    FIELD_ORDER = [
        ("Add Date", "Add Date"),
        ("Address", "Address"),
        ("AssessedValue", "Assessed Value"),
        ("Case Number", "Case Number"),
        ("Certificate Holder Name", "Certificate Holder Name"),
        ("City", "City"),
        ("Final Judgment", "Final Judgment"),
        ("My Bid", "My Bid"),
        ("Opening Bid", "Opening Bid"),
        ("Parcel ID", "Parcel ID"),
        ("PlaintiffMaxBid", "Plaintiff Max Bid"),
        ("Sale Date", "Sale Date"),
        ("Status", "Status"),
        ("Zip", "Zip"),
    ]
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{county.title()} County {sales_type.title()} Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 2em; }}
        .report-scroll-container {{
            max-width: 1700px;
            height: 1200px;
            overflow: auto;
            border: 1px solid #ccc;
            border-radius: 8px;
            background: #fff;
        }}
        .sticky-title {{
            position: sticky;
            top: 0;
            background: #fff;
            z-index: 100;
            padding-bottom: 0.5em;
            border-bottom: 2px solid #eee;
        }}
        .sticky-table-header th {{
            position: sticky;
            top: 2.2em;
            background: #f4f4f4;
            z-index: 99;
        }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
        tr:nth-child(even) {{ background: #fafafa; }}
        .filter-note {{ color: #666; font-size: 0.95em; margin-bottom: 1em; }}
    </style>
</head>
<body>
    <div class="report-scroll-container">
        <div class="sticky-title"><strong>{county.title()} County {sales_type.title()} Report</strong><br><span style="font-weight: normal; font-size: 0.95em;">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</span></div>
        <table>
            <thead class="sticky-table-header">
                <tr>'''
    html += '<th style="width:36px"><input type="checkbox" id="header-show-selected" title="Show Selected Only" style="transform: scale(1.3); cursor: pointer; vertical-align: middle;" /></th>'  # Checkbox column
    for field, label in FIELD_ORDER:
        html += f'<th>{label}</th>'
    html += '</tr>\n        </thead>\n        <tbody>\n'
    def format_currency(val):
        try:
            num = float(str(val).replace(',', '').replace('$', ''))
            if num == 0:
                return '0'
            return f"${num:,.2f}"
        except Exception:
            return val

    for row in sales:
        html += '<tr><td></td>'  # Always prepend checkbox column
        for field, _ in FIELD_ORDER:
            cell = row.get(field, "")
            if field == "Address":
                address = str(row.get("Address", "")).strip()
                city = str(row.get("City", "")).strip()
                zip_code = str(row.get("Zip", "")).strip()
                if address:
                    q = address
                    if city:
                        q += f", {city}"
                    if zip_code:
                        q += f" {zip_code}"
                    maps_url = f"https://www.google.com/maps/search/?api=1&query={q.replace(' ', '+')}"
                    html += f'<td><a href="{maps_url}" target="_blank" rel="noopener noreferrer">{address}</a></td>'
                else:
                    html += '<td></td>'
            elif field == "Parcel ID":
                parcel_id = str(cell).strip()
                # Only hyperlink if not blank, not 'TIMESHARE', and does not contain 'MULTIPLE PARCELS' (case-insensitive)
                if pa_template and parcel_id and parcel_id.upper() != 'TIMESHARE' and 'MULTIPLE PARCELS' not in parcel_id.upper() and 'TIMESHARE' not in parcel_id.upper():
                    pa_url = pa_template.replace('<<PID>>', parcel_id)
                    html += f'<td><a href="{pa_url}" target="_blank" rel="noopener noreferrer">{parcel_id}</a></td>'
                else:
                    html += f'<td>{parcel_id}</td>'
            elif field in ("Final Judgment", "Opening Bid", "AssessedValue"):
                html += f'<td>{format_currency(cell)}</td>'
            else:
                html += f'<td>{cell}</td>'
        html += '</tr>\n'
    html += '        </tbody>\n      </table>\n    </div>'
    html += '\n<script src="/sortable-table.js"></script>'
    html += '\n</body>\n</html>'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Report generated: {output_path}')

if __name__ == "__main__":
    print("=== auction_utils.py is running ===")
    json_path = 'backend/Legacy/foreclosureSales_clean.json'
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for row in data:
        county = row.get('County', '').strip()
        for sale_type in ['Foreclosure', 'Tax Deed']:
            if sale_type == 'Foreclosure':
                flag = str(row.get('UiPath', '')).strip().lower() == 'x'
                sale_list_url = row.get('List')
            else:
                flag = str(row.get('UiPathTD', '')).strip().lower() == 'x'
                sale_list_url = row.get('TaxDeedList')
            if flag and sale_list_url:
                print(county)
                print(sale_type)
                print(sale_list_url)
                exit(0)
    print('No valid county/sale type/path found.')