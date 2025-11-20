import firebase_admin
from firebase_admin import credentials, firestore
import os

# Path to your Firebase service account key
SERVICE_ACCOUNT_PATH = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'serviceAccountKey.json')

# Initialize Firebase app
if not firebase_admin._apps:
    cred = credentials.Certificate(SERVICE_ACCOUNT_PATH)
    firebase_admin.initialize_app(cred)

db = firestore.client()

# Reference to the Auctions collection
auctions_ref = db.collection('Auctions')

# Get all documents in the Auctions collection
auctions = auctions_ref.stream()

doc_count = 0
for auction in auctions:
    # Delete each document (and all subcollections)
    print(f"Deleting {auction.id}")
    auctions_ref.document(auction.id).delete()
    doc_count += 1

print(f"All {doc_count} Auctions documents deleted.")
