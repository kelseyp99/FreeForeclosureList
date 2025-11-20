import os
from google.cloud import firestore
from google.oauth2 import service_account

# Path to your service account key
SERVICE_ACCOUNT_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "/Users/tinman/Projects/FreeForeclosureList/backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json")

credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH)
db = firestore.Client(credentials=credentials)

def delete_all_sales():
    sales_ref = db.collection("Sales")
    docs = list(sales_ref.stream())
    print(f"Found {len(docs)} documents in Sales collection.")
    for i, doc in enumerate(docs, 1):
        print(f"Deleting {i}/{len(docs)}: {doc.id}")
        doc.reference.delete()
    print("All Sales documents deleted.")

if __name__ == "__main__":
    delete_all_sales()
