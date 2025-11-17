import csv
import json

# Input and output paths
csv_path = 'Legacy/foreclosureSales.csv'
json_path = 'Legacy/foreclosureSales.json'

# Read CSV and convert to list of dicts
with open(csv_path, 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    rows = [row for row in reader if any(row.values())]

# Write as JSON array
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(rows, f, indent=2)

print(f'Converted {csv_path} to {json_path} (array of {len(rows)} objects)')
