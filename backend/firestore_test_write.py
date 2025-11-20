from google.cloud import firestore
from google.oauth2 import service_account

# Path to your service account key
SERVICE_ACCOUNT_PATH = "/Users/tinman/Projects/FreeForeclosureList/backend/foreclosure-15f09-firebase-adminsdk-fbsvc-0fd54751e3.json"

credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH)
db = firestore.Client(credentials=credentials)

db.collection("TestCollection").document("TestDoc").set({"hello": "world"})
print("Wrote test doc.")
