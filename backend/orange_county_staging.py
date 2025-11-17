import os
import csv
import datetime
import firebase_admin
from firebase_admin import credentials, firestore

# Initialize Firebase (update the path and project info as needed)
FIREBASE_CRED_PATH = 'path/to/serviceAccountKey.json'  # TODO: Set this to your actual key file
FIREBASE_PROJECT_ID = 'your-firebase-project-id'       # TODO: Set this to your actual project id

if not firebase_admin._apps:
    cred = credentials.Certificate(FIREBASE_CRED_PATH)
    firebase_admin.initialize_app(cred, {
        'projectId': FIREBASE_PROJECT_ID
    })
db = firestore.client()

def parse_date(date_str):
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y'):
        try:
            return datetime.datetime.strptime(date_str, fmt).date()
        except Exception:
            continue
    return None

def process_orange_county_csv(file_path):
    """Parse Orange County foreclosure CSV file and return list of dicts."""
    results = []
    with open(file_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        headers = next(reader, None)
        for row in reader:
            if not row or len(row) < 5:
                continue
            # Map columns as needed (update indices if your CSV differs)
            result = {
                'sale_date': parse_date(row[0]),
                'add_date': parse_date(row[1]),
                'case_number': row[2],
                'status': row[3],
                'final_judgment': float(row[4]) if row[4] else None,
                'opening_bid': float(row[5]) if len(row) > 5 and row[5] else None,
                'address': row[10] if len(row) > 10 else '',
                'city': row[11] if len(row) > 11 else '',
                'zip': row[12] if len(row) > 12 else '',
                'parcel_id': row[13] if len(row) > 13 else '',
                'county': 'Orange',
            }
            results.append(result)
    return results

def upload_results_to_firebase(county, sale_type, results):
    collection = db.collection('foreclosures').document(county).collection(sale_type)
    for result in results:
        doc_id = f"{result['sale_date']}_{result['case_number']}"
        collection.document(doc_id).set(result)

def main():
    # Example usage
    file_path = 'staging/OrangeCountySalesList.csv'  # TODO: Set to your actual file
    sale_type = 'Foreclosure'
    results = process_orange_county_csv(file_path)
    upload_results_to_firebase('Orange', sale_type, results)
    print(f"Uploaded {len(results)} records for Orange County")

if __name__ == '__main__':
    main()
