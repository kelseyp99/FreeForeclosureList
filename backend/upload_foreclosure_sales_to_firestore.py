import json
import firebase_admin
from firebase_admin import credentials, firestore

# Path to your service account key
SERVICE_ACCOUNT_PATH = 'backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json'
# Your Firestore project ID (update if needed)
PROJECT_ID = 'foreclosure-15f09'
# Path to your JSON file
JSON_PATH = 'backend/Legacy/foreclosureSales_clean.json'
# Firestore collection name
COLLECTION_NAME = 'auction_parameters'

# Initialize Firebase Admin
cred = credentials.Certificate(SERVICE_ACCOUNT_PATH)
firebase_admin.initialize_app(cred, {'projectId': PROJECT_ID})
db = firestore.client()

# Load JSON data
with open(JSON_PATH, 'r', encoding='utf-8') as f:
    data = json.load(f)

# Upload each row as a document (using County as document ID if present)
for row in data:
    doc_id = row.get('County') or None
    if doc_id:
        db.collection(COLLECTION_NAME).document(doc_id).set(row)
    else:
        db.collection(COLLECTION_NAME).add(row)

print(f"Uploaded {len(data)} documents to Firestore collection '{COLLECTION_NAME}'")
