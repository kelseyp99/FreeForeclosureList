import csv
import os
from datetime import datetime

# Path to your QuickSearch.csv file

import argparse

# Default path to your QuickSearch.csv file
DEFAULT_CSV_PATH = os.path.join(os.path.dirname(__file__), 'Legacy', 'QuickSearch.csv')
# Output HTML report path
REPORTS_DIR = os.path.join(os.path.dirname(__file__), '..', 'public', 'reports')
REPORT_FILENAME = f'sales_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html'
REPORT_PATH = os.path.join(REPORTS_DIR, REPORT_FILENAME)

# Ensure reports directory exists
os.makedirs(REPORTS_DIR, exist_ok=True)

def generate_html_report(csv_path, output_path):
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        rows = list(reader)


    html = (
        '<!DOCTYPE html>\n'
        '<html lang="en">\n'
        '<head>\n'
        '    <meta charset="UTF-8">\n'
        '    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        '    <title>Sales Report</title>\n'
        '    <style>\n'
        '        body { font-family: Arial, sans-serif; margin: 2em; }\n'
        '        table { border-collapse: collapse; width: 100%; }\n'
        '        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }\n'
        '        th { background: #f4f4f4; }\n'
        '        tr:nth-child(even) { background: #fafafa; }\n'
        '    </style>\n'
        '</head>\n'
        '<body>\n'
        f'    <h1>Sales Report</h1>\n'
        f'    <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>\n'
        '    <table>\n'
        '        <thead>\n'
        '            <tr>'
    )
    import os
    from datetime import datetime
    import argparse
    import firebase_admin
    from firebase_admin import credentials, firestore

    # Path to your Firebase service account key
    SERVICE_ACCOUNT_PATH = os.path.join(os.path.dirname(__file__), 'foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json')
    # Output HTML report path
    REPORTS_DIR = os.path.join(os.path.dirname(__file__), '..', 'public', 'reports')
    REPORT_FILENAME = f'sales_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html'
    REPORT_PATH = os.path.join(REPORTS_DIR, REPORT_FILENAME)

    # Ensure reports directory exists
    os.makedirs(REPORTS_DIR, exist_ok=True)

    def fetch_sales_from_firestore():
        # Initialize Firebase Admin
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

    def generate_html_report_from_sales(sales, output_path):
        html = (
            '<!DOCTYPE html>\n'
            '<html lang="en">\n'
            '<head>\n'
            '    <meta charset="UTF-8">\n'
            '    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
            '    <title>Sales Report</title>\n'
            '    <style>\n'
            '        body { font-family: Arial, sans-serif; margin: 2em; }\n'
            '        table { border-collapse: collapse; width: 100%; }\n'
            '        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }\n'
            '        th { background: #f4f4f4; }\n'
            '        tr:nth-child(even) { background: #fafafa; }\n'
            '    </style>\n'
            '</head>\n'
            '<body>\n'
            f'    <h1>Sales Report</h1>\n'
            f'    <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>\n'
            '    <table>\n'
            '        <thead>\n'
            '            <tr>'
        )
        # Table headers
        if sales:
            for field in sales[0].keys():
                html += f'<th>{field}</th>'
        html += '</tr>\n        </thead>\n        <tbody>\n'
        # Table rows
        for row in sales:
            html += '<tr>'
            for value in row.values():
                html += f'<td>{value}</td>'
            html += '</tr>\n'
        html += '        </tbody>\n    </table>\n</body>\n</html>'

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f'Report generated: {output_path}')

    if __name__ == '__main__':
        parser = argparse.ArgumentParser(description='Generate an HTML sales report from Firestore.')
        parser.add_argument('--output', '-o', type=str, default=REPORT_PATH, help='Path to the output HTML file (default: auto-named in public/reports/)')
        args = parser.parse_args()
        sales = fetch_sales_from_firestore()
        generate_html_report_from_sales(sales, args.output)
