import json
from datetime import datetime
import sys

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

def get_next_county_and_sale_type():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    soonest = None
    soonest_type = None
    soonest_county = None
    soonest_date = None
    for row in data:
        county = row.get('County', '').strip()
        fc_flag = str(row.get('UiPath', '')).strip().lower() == 'x'
        td_flag = str(row.get('UiPathTD', '')).strip().lower() == 'x'
        fc_date = parse_date(row.get('Foreclosure', '')) if fc_flag else None
        td_date = parse_date(row.get('Tax Deed', '')) if td_flag else None
        for sale_type, date in [('Foreclosure', fc_date), ('Tax Deed', td_date)]:
            if date:
                if soonest_date is None or date < soonest_date:
                    soonest_date = date
                    soonest_type = sale_type
                    soonest_county = county
    if soonest_county and soonest_type:
        print(f"{soonest_county} {soonest_type}")
    else:
        print("")

if __name__ == "__main__":
    get_next_county_and_sale_type()
