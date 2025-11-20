import csv

import random
from datetime import datetime
import firebase_admin
from firebase_admin import credentials, firestore
import os
import sys

# Path to your Firebase service account key
SERVICE_ACCOUNT_PATH = 'backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json'
# Default CSV path, can be overridden by CLI arg
CSV_PATH = sys.argv[1] if len(sys.argv) > 1 else 'backend/Legacy/QuickSearch.csv'

# Initialize Firebase Admin
if not firebase_admin._apps:
    cred = credentials.Certificate(SERVICE_ACCOUNT_PATH)
    firebase_admin.initialize_app(cred)
db = firestore.client()

# Map CSV columns to Firestore fields (add more as needed)
FIELD_MAP = {
    'Sale Date': 'Sale Date',
    'Add Date': 'Add Date',
    'Case Number': 'Case Number',
    'Status': 'Status',
    'Final Judgment': 'Final Judgment',
    'Opening Bid': 'Opening Bid',
    'Assessed Value': 'AssessedValue',
    'Certificate Holder Name': 'Plaintiff Name',
    'Plaintiff Max Bid': 'PlaintiffMaxBid',
    'Address': 'Address',
    'City': 'City',
    'Zip': 'Zip',
    'Parcel ID': 'PID',
    'My Bid': 'My Bid',
}



# Infer sales type from filename
def infer_sales_type_from_filename(filename):
    fname = os.path.basename(filename).lower()
    if 'foreclosure' in fname:
        return 'Foreclosure'
    elif 'taxdeed' in fname or 'tax_deed' in fname or 'tax-deed' in fname:
        return 'Tax Deed'
    else:
        return 'Unknown'

# Infer county from filename (before first dash or underscore)
def infer_county_from_filename(filename):
    fname = os.path.basename(filename)
    # Split on dash or underscore, take first part
    for sep in ['-', '_']:
        if sep in fname:
            return fname.split(sep)[0].strip().title()
    return 'Unknown'

SALES_TYPE = infer_sales_type_from_filename(CSV_PATH)
COUNTY = infer_county_from_filename(CSV_PATH)

MOCK_FIELDS = {
    'Sales Type': lambda row: SALES_TYPE,
    'notes': lambda row: 'Mock processed',
    'Timestamp': lambda row: datetime.now().isoformat(),
    'Telegram': lambda row: f'https://t.me/mock/{random.randint(1000,9999)}',
    'Zillo': lambda row: f'https://zillow.com/homedetails/{random.randint(100000,999999)}',
    'Realtor.com': lambda row: f'https://realtor.com/realestateandhomes-detail/{random.randint(100000,999999)}',
    'RedFin': lambda row: f'https://redfin.com/FL/{row.get("City", "")}/{random.randint(100000,999999)}',
}

def process_and_upload_sales():
    print(f"Processing file: {CSV_PATH}")
    print(f"Inferred Sales Type: {SALES_TYPE}")
    print(f"Inferred County: {COUNTY}")
    with open(CSV_PATH, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            sale = {}
            for csv_col, fs_field in FIELD_MAP.items():
                sale[fs_field] = row.get(csv_col, '').strip()
            # Add mock/demo fields
            for k, v in MOCK_FIELDS.items():
                sale[k] = v(row)
            # Add inferred county
            sale['County'] = COUNTY
            # Add any other fields you want to mock here
            print(f"Uploading sale: {sale['Case Number']} ({sale['Address']}) | County: {sale['County']} | Sales Type: {sale['Sales Type']}")
            db.collection('sales').add(sale)
    print("All sales uploaded.")

if __name__ == '__main__':
    process_and_upload_sales()
