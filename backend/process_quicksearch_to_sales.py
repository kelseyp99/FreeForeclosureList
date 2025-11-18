import csv
import random
from datetime import datetime
import firebase_admin
from firebase_admin import credentials, firestore

# Path to your Firebase service account key
SERVICE_ACCOUNT_PATH = 'backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json'
CSV_PATH = 'backend/Legacy/QuickSearch.csv'

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

# Add mock fields for demonstration
MOCK_FIELDS = {
    'Sales Type': lambda row: 'Foreclosure',
    'notes': lambda row: 'Mock processed',
    'Timestamp': lambda row: datetime.now().isoformat(),
    'Telegram': lambda row: f'https://t.me/mock/{random.randint(1000,9999)}',
    'Zillo': lambda row: f'https://zillow.com/homedetails/{random.randint(100000,999999)}',
    'Realtor.com': lambda row: f'https://realtor.com/realestateandhomes-detail/{random.randint(100000,999999)}',
    'RedFin': lambda row: f'https://redfin.com/FL/{row.get("City", "")}/{random.randint(100000,999999)}',
}

def process_and_upload_sales():
    with open(CSV_PATH, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            sale = {}
            for csv_col, fs_field in FIELD_MAP.items():
                sale[fs_field] = row.get(csv_col, '').strip()
            # Add mock/demo fields
            for k, v in MOCK_FIELDS.items():
                sale[k] = v(row)
            # Add any other fields you want to mock here
            print(f"Uploading sale: {sale['Case Number']} ({sale['Address']})")
            db.collection('sales').add(sale)
    print("All sales uploaded.")

if __name__ == '__main__':
    process_and_upload_sales()
