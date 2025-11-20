import os
import sys
import csv
import re
from google.cloud import firestore
from google.oauth2 import service_account

# --- CONFIG ---
# Path to your Firebase service account key
SERVICE_ACCOUNT_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "serviceAccountKey.json")

# --- FIRESTORE INIT ---
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_PATH)
db = firestore.Client(credentials=credentials)

# --- HELPERS ---

# Infer sales type from filename
def infer_sales_type_from_filename(filename):
    fname = os.path.basename(filename).lower()
    if 'foreclosure' in fname:
        return 'foreclosure'
    elif 'taxdeed' in fname or 'tax_deed' in fname or 'tax-deed' in fname:
        return 'taxdeed'
    else:
        return 'unknown'

def parse_filename(filename):
    # Example: Orange-20251119_Foreclosure_QuickSearch.csv
    base = os.path.basename(filename)
    print(f"Parsing filename: {base}")  # Debug print
    m = re.match(r"([A-Za-z]+)[-_](\d{8})[_-]?(Foreclosure|TaxDeed)?[_-]?QuickSearch\.csv", base, re.IGNORECASE)
    if not m:
        raise ValueError(f"Filename {filename} does not match expected pattern.")
    county, date, sale_type = m.group(1), m.group(2), m.group(3)
    # If sale_type is None, try to infer from filename
    if not sale_type:
        sale_type = infer_sales_type_from_filename(filename)
    else:
        sale_type = sale_type.lower()
    return county, date, sale_type

def upsert_sale(county, date, sale_type, row):
    casenumber = row.get("Case Number") or row.get("CaseNumber") or row.get("CaseNo")
    if not casenumber:
        print(f"Skipping row with missing case number: {row}")
        return
    doc_name = f"{county}_{date}"
    if sale_type and sale_type != 'unknown':
        doc_name += f"_{sale_type}"
    doc_ref = db.collection("Auctions").document(doc_name) \
        .collection("Sales").document(str(casenumber))
    print(f"Writing to: Auctions/{doc_name}/Sales/{casenumber}")
    print(f"Data: {row}")
    existing = doc_ref.get()
    if existing.exists:
        # Only update if non-PK fields changed
        old = existing.to_dict()
        update = {k: v for k, v in row.items() if k not in ("Case Number", "CaseNumber", "CaseNo") and old.get(k) != v}
        if update:
            doc_ref.update(update)
            print(f"Updated sale: {casenumber}")
    else:
        doc_ref.set(row)
        print(f"Inserted sale: {casenumber}")

def main(csv_path):
    county, date, sale_type = parse_filename(csv_path)
    total = 0
    written = 0
    skipped = 0
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        # Normalize fieldnames to strip whitespace
        reader.fieldnames = [fn.strip() for fn in reader.fieldnames]
        for row in reader:
            # Also strip whitespace from keys in each row
            row = {k.strip(): v for k, v in row.items()}
            total += 1
            casenumber = row.get("Case Number") or row.get("CaseNumber") or row.get("CaseNo")
            if not casenumber:
                skipped += 1
                print(f"Skipping row {total}: missing case number")
                continue
            written += 1
            upsert_sale(county, date, sale_type, row)
    print(f"\nSummary for {csv_path}:")
    print(f"  Total rows: {total}")
    print(f"  Written: {written}")
    print(f"  Skipped (missing case number): {skipped}")
    print(f"Done processing {csv_path} for {county} on {date}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python process_quicksearch_to_auctions.py <QuickSearch.csv>")
        sys.exit(1)
    main(sys.argv[1])
