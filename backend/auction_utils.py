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

def process_quicksearch_to_auctions( county, sale_type, csv_path):
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


def get_next_sale(min_hours=0):
    """
    Loops through auction_parameters to find the first blank (missing) Foreclosure or Tax Deed last update,
    or, if none are blank, returns the one with the oldest timestamp.
    Returns (county, sale_type, url).
    """
    from datetime import datetime, timezone

    params_ref = db.collection("auction_parameters")
    docs = list(params_ref.stream())
    now = datetime.now(timezone.utc)

    # Helper: decide whether we should process this county/sale_type
    def should_process(data, sale_type):
        if sale_type.lower() == "foreclosure":
            return data.get("processForeclosure", data.get("ForeclosureInclude", True))
        if sale_type.lower() == "tax deed":
            return data.get("processTaxDeed", data.get("TaxDeedInclude", True))
        return True

    # First loop: find blank last update for a sale type that is configured to be processed
    for doc in docs:
        data = doc.to_dict()
        county = doc.id
        fc_time = data.get("ForeclosureLastUpdate")
        td_time = data.get("TaxDeedLastUpdate")
        foreclosure_url = data.get("List")
        taxdeed_url = data.get("TaxDeedList")

        if should_process(data, "Foreclosure") and not fc_time and foreclosure_url:
            return county, "Foreclosure", foreclosure_url
        if should_process(data, "Tax Deed") and not td_time and taxdeed_url:
            return county, "Tax Deed", taxdeed_url

    # Second loop: find oldest last_update among eligible sale types
    oldest_candidate = None
    oldest_time = None
    for doc in docs:
        data = doc.to_dict()
        county = doc.id
        fc_time = data.get("ForeclosureLastUpdate")
        td_time = data.get("TaxDeedLastUpdate")
        foreclosure_url = data.get("List")
        taxdeed_url = data.get("TaxDeedList")

        for sale_type, last_update, url in [
            ("Foreclosure", fc_time, foreclosure_url),
            ("Tax Deed", td_time, taxdeed_url)
        ]:
            # Skip sale types that are not configured to be processed
            if not should_process(data, sale_type):
                continue
            if not last_update or not url:
                continue
            if hasattr(last_update, 'replace'):
                last_update_dt = last_update.replace(tzinfo=timezone.utc)
            else:
                last_update_dt = datetime.fromisoformat(str(last_update))
            if (oldest_time is None) or (last_update_dt < oldest_time):
                oldest_time = last_update_dt
                oldest_candidate = (county, sale_type, url)

    if oldest_candidate and oldest_candidate[2]:
        return oldest_candidate
    return None, None, None

def upload_report_and_mark_updated(report_path, county, sale_type, dest_dir='dist/reports'):
    """
    Ensures the report HTML exists, copies it to public/reports for deployment,
    updates Firestore auction_parameters last_update, and deploys to Firebase Hosting.
    """
    import shutil
    import datetime
    import os
    import subprocess

    # 1. Always copy to public/reports (for deploy)
    public_reports_dir = os.path.join(os.path.dirname(__file__), '..', 'public', 'reports')
    if not os.path.exists(public_reports_dir):
        os.makedirs(public_reports_dir)
    public_dest_path = os.path.join(public_reports_dir, os.path.basename(report_path))
    try:
        if os.path.abspath(report_path) != os.path.abspath(public_dest_path):
            shutil.copy2(report_path, public_dest_path)
            print(f"[INFO] Copied {report_path} to {public_dest_path}")
        else:
            print(f"[INFO] Source and destination are the same file ({report_path}), skipping copy.")
    except Exception as e:
        print(f"[ERROR] Could not copy {report_path} to {public_dest_path}: {e}")

    # 2. (Optional) Run the build step if you want to update frontend code
    # To avoid unnecessary builds, comment out or remove the build step unless needed
    build_output = ""
    build_error = ""

    # 3. Update Firestore last update for the correct sale type field
    county_doc = db.collection("auction_parameters").document(county)
    now = datetime.datetime.utcnow()
    if sale_type.lower() == "foreclosure":
        update_data = {"ForeclosureLastUpdate": now}
    elif sale_type.lower() == "tax deed":
        update_data = {"TaxDeedLastUpdate": now}
    else:
        update_data = {f"{sale_type}LastUpdate": now}
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
        "copied_to": public_dest_path,
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

def generate_html_report_from_firestore(county, sales_type):
    # Uses the global db object (already created at the top of your file
    sales_ref = db.collection('sales')
    docs = sales_ref.stream()
    sales = [doc.to_dict() for doc in docs]
    filtered = filter_sales(sales, county, sales_type)
    return generate_html_report_from_sales(filtered, county, sales_type)

def generate_html_report_from_sales(sales, county, sales_type):
    import os
    from datetime import datetime
    output_path = f"dist/reports/sales_report_{county.lower()}_{sales_type.lower().replace(" ", "")}.html"
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
    return output_path

def filter_sales(sales, county, sales_type):
    filtered = []
    print(f"[DEBUG] Filtering sales for county='{county}', sales_type='{sales_type}'")
    for row in sales:
        row_county = (row.get('County') or '').strip().lower()
        row_type = (row.get('SaleType') or row.get('Sales Type') or '').strip().lower()
        if row_county == county.lower() and row_type == sales_type.lower():
            filtered.append(row)
    print(f"[DEBUG] Total matches found: {len(filtered)}")
    return filtered

def testAll():
    print("=== auction_utils.py is running ===")
    csv_path = f"/Users/tinman/Downloads/QuickSearch.csv"
    min_hours = 24

    # Before returning, always delete the existing QuickSearch.csv to avoid suffixing in automation
    quicksearch_path = csv_path
    import os
    if os.path.exists(quicksearch_path):
        try:
            os.remove(quicksearch_path)
            print(f"[INFO] Deleted existing {quicksearch_path} to avoid suffixing (from get_next_sale).")
        except Exception as e:
            print(f"[WARN] Could not delete {quicksearch_path}: {e}")

    #get the next sale to process
    county, sale_type, sale_list_url = get_next_sale(min_hours=min_hours)
    print(county)
    print(sale_type)
    print(sale_list_url)
    user_input = input("Press Enter to continue or x to exclude moving forward...")

    if user_input.strip().lower() == 'x':
        mark_county_sale_type_excluded(county, sale_type)
        print(f"[INFO] Excluded {county} {sale_type} from future processing. Skipping workflow steps.")
        return
    # --- Begin workflow steps ---
    result = process_quicksearch_to_auctions(county, sale_type, csv_path)
    print(f"Processed QuickSearch: {result}")
    # Generate report from Firestore
    print(f"[DEBUG] Generating report for county='{county}', sale_type='{sale_type}'")
    output_path = generate_html_report_from_firestore(county, sale_type)
    print(f"[DEBUG] Generated report: {output_path}")
    # Upload report and deploy
    upload_result = upload_report_and_mark_updated(output_path, county, sale_type)
    print(f"Upload and deploy result: {upload_result}")
    exit(0)

def mark_county_sale_type_excluded(county, sale_type):
    """
    Sets processForeclosure or processTaxDeed to False for the given county and sale_type in auction_parameters.
    """
    field = None
    if sale_type.lower() == "foreclosure":
        field = "processForeclosure"
    elif sale_type.lower() == "tax deed":
        field = "processTaxDeed"
    else:
        print(f"[WARN] Unknown sale_type: {sale_type}")
        return
    doc_ref = db.collection("auction_parameters").document(county)
    doc_ref.set({field: False}, merge=True)
    print(f"[INFO] Marked {county} {sale_type} as excluded (set {field}=False)")

import sys
def arg(n):
    try:
        return sys.argv[n]
    except IndexError:
        return None

if __name__ == "__main__":
    if arg(1) == "generate_html_report_from_firestore":
        county = arg(2)
        sales_type = arg(3)
        print(generate_html_report_from_firestore(county, sales_type))
    elif arg(1) == "upload_report_and_mark_updated":
        county = arg(2)
        sales_type = arg(3)
        output_path = arg(4)
        print(upload_report_and_mark_updated(output_path, county, sales_type))
    elif arg(1) == "testAll":
        testAll()
    else:
        # Default action if no or unknown argument is given
        testAll()