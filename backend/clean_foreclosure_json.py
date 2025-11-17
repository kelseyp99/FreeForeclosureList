import json

# Input/output paths
input_path = 'backend/Legacy/foreclosureSales.json'
output_path = 'backend/Legacy/foreclosureSales_clean.json'

with open(input_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

cleaned = []
for row in data:
    # Remove rows with no County or that are the 'Start at' row
    if not row.get('County') or row['County'].strip().lower() == 'start at':
        continue
    # Remove keys that are empty or not strings
    clean_row = {str(k).strip(): v for k, v in row.items() if k and str(k).strip()}
    cleaned.append(clean_row)

with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(cleaned, f, indent=2)

print(f'Cleaned {len(cleaned)} records written to {output_path}')
