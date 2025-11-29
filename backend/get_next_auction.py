import json
from datetime import datetime
import sys

JSON_PATH = 'backend/Legacy/foreclosureSales_clean.json'

COLUMN_ORDER = [
    "County",
    "Days FC",
    "Days TD",
    "Foreclosure",
    "List",
    "Court Docs",
    "PA",
    "PA template",
    "PA template Address",
    "Sale Template",
    "Tax Deed",
    "TaxDeedList",
    "UiPath",
    "UiPathTD",
    "WIXparamID",
    "download file",
    "public records",
    "updateWIX"
]

def parse_date(date_str):
    try:
        return datetime.strptime(date_str.strip(), '%m/%d/%Y')
    except Exception:
        return None

def get_next_sale_for_county(county_name):
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
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

def generate_sales_report_for_county(county, sales_type):
    """
    Generate and update the HTML report for the given county and sales type.
    Usage: generate_sales_report_for_county('Orange', 'Foreclosure')
    """
    import os
    import json
    from datetime import datetime
    REPORTS_DIR = os.path.join(os.path.dirname(__file__), '..', 'public', 'reports')
    os.makedirs(REPORTS_DIR, exist_ok=True)
    # Firestore fetch
    SERVICE_ACCOUNT_PATH = "/Users/tinman/Projects/FreeForeclosureList/backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json"
    from google.cloud import firestore
    from google.oauth2 import service_account
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH)
    db = firestore.Client(credentials=credentials)
    sales_ref = db.collection('sales')
    docs = sales_ref.stream()
    sales = [doc.to_dict() for doc in docs]
    # Filter
    filtered = []
    for row in sales:
        row_county = (row.get('County') or '').strip().lower()
        row_type = (row.get('Sales Type') or '').strip().lower()
        if row_county == county.lower() and row_type == sales_type.lower():
            filtered.append(row)
    if not filtered:
        print(f"No sales found for county '{county}' and sales type '{sales_type}'.")
        return False
    # PA template
    pa_template = None
    pa_template_path = os.path.join(os.path.dirname(__file__), 'Legacy', 'foreclosureSales_clean.json')
    try:
        with open(pa_template_path, 'r', encoding='utf-8') as f:
            pa_data = json.load(f)
            for entry in pa_data:
                if entry.get('County', '').strip().lower() == county.strip().lower():
                    pa_template = entry.get('PA template', '').strip()
                    break
    except Exception:
        pa_template = None
    # Field order
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
    # HTML
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
    # Removed 'Show Selected Only' checkbox from table header
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
    for row in filtered:
        html += '<tr><td></td>'
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
    output_path = os.path.join(REPORTS_DIR, f"sales_report_{county.lower()}_{sales_type.lower().replace(' ', '')}.html")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Report generated: {output_path}')
    return output_path

def get_next_county_and_sale_type():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    soonest = None
    soonest_type = None
    soonest_county = None
    soonest_date = None
    for row in data:
        county = row.get('County', '').strip()
        fc_flag = str(row.get('UiPath', '')).strip().lower() == 'x'
        td_flag = str(row.get('UiPathTD', '')).strip().lower() == 'x'
        fc_date = parse_date(row.get('Foreclosure', '')) if fc_flag else None
        td_date = parse_date(row.get('Tax Deed', '')) if td_flag else None
        for sale_type, date in [('Foreclosure', fc_date), ('Tax Deed', td_date)]:
            if date:
                if soonest_date is None or date < soonest_date:
                    soonest_date = date
                    soonest_type = sale_type
                    soonest_county = county
    if soonest_county and soonest_type:
        print(f"{soonest_county} {soonest_type}")
    else:
        print("")

if __name__ == "__main__":
    get_next_county_and_sale_type()
