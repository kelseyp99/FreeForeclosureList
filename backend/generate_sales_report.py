
import csv
import os
from datetime import datetime
import argparse

REPORTS_DIR = os.path.join(os.path.dirname(__file__), '..', 'public', 'reports')
os.makedirs(REPORTS_DIR, exist_ok=True)

def fetch_sales_from_firestore():
    import firebase_admin
    from firebase_admin import credentials, firestore
    SERVICE_ACCOUNT_PATH = os.path.join(os.path.dirname(__file__), 'foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json')
    if not firebase_admin._apps:
        cred = credentials.Certificate(SERVICE_ACCOUNT_PATH)
        firebase_admin.initialize_app(cred)
    db = firestore.client()
    sales_ref = db.collection('sales')
    docs = sales_ref.stream()
    sales = []
    for doc in docs:
        sales.append(doc.to_dict())
    return sales

def filter_sales(sales, county, sales_type):
    filtered = []
    for row in sales:
        row_county = (row.get('County') or '').strip().lower()
        row_type = (row.get('Sales Type') or '').strip().lower()
        parcel_id = (row.get('Parcel ID') or '').strip()
        is_timeshare = (row.get('Certificate Holder Name') or '').strip().upper() == 'TIMESHARE' or (row.get('Parcel ID') or '').strip().upper() == 'TIMESHARE'
        if row_county == county.lower() and row_type == sales_type.lower():
            if is_timeshare:
                continue
            if not parcel_id:
                continue
            filtered.append(row)
    return filtered

def generate_html_report_from_sales(sales, output_path, county, sales_type):
    # Define the desired field order and labels
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
      <div class="filter-note">Filtered: Timeshares and rows with blank Parcel ID are excluded.</div>
      <table>
        <thead class="sticky-table-header">
            <tr>'''
    html += '<th style="width:36px">★</th>'  # Checkbox column
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
            elif field in ("Final Judgment", "Opening Bid", "AssessedValue"):
                html += f'<td>{format_currency(cell)}</td>'
            else:
                html += f'<td>{cell}</td>'
        html += '</tr>\n'
    html += '        </tbody>\n      </table>\n    </div>\n</body>\n</html>'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Report generated: {output_path}')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate a county/sales type HTML sales report from Firestore or CSV.')
    parser.add_argument('--county', required=True, help='County name (e.g., Orange)')
    parser.add_argument('--sales_type', required=True, help='Sales type (e.g., Foreclosure, Tax Deed)')
    parser.add_argument('--csv', help='Optional path to CSV file. If not provided, pulls from Firestore.')
    args = parser.parse_args()

    if args.csv:
        with open(args.csv, newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            sales = list(reader)
    else:
        sales = fetch_sales_from_firestore()

    filtered = filter_sales(sales, args.county, args.sales_type)
    if not filtered:
        print(f"No sales found for county '{args.county}' and sales type '{args.sales_type}'.")
    else:
        output_path = os.path.join(REPORTS_DIR, f"sales_report_{args.county.lower()}_{args.sales_type.lower().replace(' ', '')}.html")
        generate_html_report_from_sales(filtered, output_path, args.county, args.sales_type)
