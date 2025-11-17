import csv
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

def upload_csv_to_firestore(csv_path, collection_name):
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Use a unique field or auto-generated ID for each document
            doc_id = row.get('id') or row.get('county') or None
            if doc_id:
                db.collection(collection_name).document(doc_id).set(row)
            else:
                db.collection(collection_name).add(row)
    print(f"Uploaded CSV to Firestore collection '{collection_name}'")

if __name__ == '__main__':
    # Example usage
    csv_path = 'path/to/your/parameter_table.csv'  # TODO: Set to your CSV file
    collection_name = 'auction_parameters'         # Name for your Firestore collection
    upload_csv_to_firestore(csv_path, collection_name)
