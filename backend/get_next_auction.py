
import json
from datetime import datetime

JSON_PATH = 'backend/Legacy/foreclosureSales_clean.json'

COLUMN_ORDER = [
    "County",
    "Days FC",
    "Days TD",
    "Foreclosure",
    "List",
    "Court Docs",
    "PA",
    "PA template",
    "PA template Address",
    "Sale Template",
    "Tax Deed",
    "TaxDeedList",
    "UiPath",
    "UiPathTD",
    "WIXparamID",
    "download file",
    "public records",
    "updateWIX"
]

def parse_date(date_str):
    try:
        return datetime.strptime(date_str.strip(), '%m/%d/%Y')
    except Exception:
        return None

def get_next_sale_for_county(county_name):
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for row in data:
        if row.get('County', '').strip().lower() == county_name.strip().lower():
            fc_flag = str(row.get('UiPath', '')).strip().lower() == 'x'
            td_flag = str(row.get('UiPathTD', '')).strip().lower() == 'x'
            fc_date = parse_date(row.get('Foreclosure', '')) if fc_flag else None
            td_date = parse_date(row.get('Tax Deed', '')) if td_flag else None
            if fc_date and td_date:
                if fc_date <= td_date:
                    return 'Foreclosure', fc_date.strftime('%m/%d/%Y'), row
                else:
                    return 'Tax Deed', td_date.strftime('%m/%d/%Y'), row
            elif fc_date:
                return 'Foreclosure', fc_date.strftime('%m/%d/%Y'), row
            elif td_date:
                return 'Tax Deed', td_date.strftime('%m/%d/%Y'), row
            else:
                return None, None, row
    return None, None, None

if __name__ == "__main__":
    county = 'Pasco'
    sale_type, sale_date, row = get_next_sale_for_county(county)
    if sale_type:
        print(f"Next sale for {county}: {sale_type} on {sale_date}")
        for col in COLUMN_ORDER:
            print(f"{col}: {row.get(col, '')}")
    else:
        print(f"No upcoming sale found for {county}.")
