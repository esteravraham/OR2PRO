#סקריפט שמבצע משיכת נתונים מה-API וההשוואה לcleaned_Database

import requests
import pandas as pd
import json

# ייבוא הפונקציות מהמודול שיצרנו
from data_utils import extract_phone, api_db_match

INPUT_FILE = "Cleaned_Database.xlsx"
OUTPUT_FILE = "Ramat_Gan_API_Match_Report.xlsx"
API_URL = "https://api-m.ramat-gan.muni.il/api/MyNeighborhood/object/he?c=57206&c=62612&c=62613"

try:
    df = pd.read_excel(INPUT_FILE, sheet_name="Cleaned_Data")

    print("Fetching data from API...")
    response = requests.get(API_URL, timeout=20)
    response.raise_for_status()
    api_data = json.loads(response.text).get("locations", [])

    matched_rows = []
    unmatched_api_rows = []
    matched_db_indices = set()

    print("Cross-referencing data...")

    for item in api_data:
        api_address_raw = item.get("address", "")
        api_phone = extract_phone(item.get("phone", ""))
        api_name = item.get("name", "")
        api_neighborhood = item.get("neighborhood", "")
        
        geom = item.get("geometricObject", {})
        api_x = geom.get("x", None)
        api_y = geom.get("y", None)

        is_match = False

        for idx, row in df.iterrows():
            db_city = str(row.get('city', 'Unknown')).strip()
            db_addr_raw = str(row.get('address', 'Unknown')).strip()
            full_db_address = f"{db_addr_raw} {db_city}"
            
            db_phone = extract_phone(row.get('phone', ''))
            db_name = row.get('name', '')

            # שימוש בפונקציות המיובאות
            if api_db_match(api_address_raw, full_db_address, api_phone, db_phone, api_name, db_name):
                    matched_db_indices.add(idx)
                    matched_rows.append({
                        'id': row.get('id', ''),
                        'db_name': db_name,
                        'api_name': api_name,
                        'city': db_city,
                        'db_address': db_addr_raw,
                        'api_address': api_address_raw,
                        'neighborhood': api_neighborhood,
                        'x': api_x,
                        'y': api_y,
                        'ownership': row.get('ownership', ''),
                        'manager': row.get('manager', '')
                    })
                    is_match = True
                    break
        
        if not is_match:
            unmatched_api_rows.append({
                'api_name': api_name,
                'api_address': api_address_raw,
                'neighborhood': api_neighborhood,
                'x': api_x,
                'y': api_y
            })

    unmatched_db_rows = []
    for idx, row in df.iterrows():
        if idx not in matched_db_indices:
            unmatched_db_rows.append({
                'id': row.get('id', ''),
                'name': row.get('name', ''),
                'city': row.get('city', ''),
                'address': row.get('address', ''),
                'ownership': row.get('ownership', ''),
                'manager': row.get('manager', '')
            })

    matched_df = pd.DataFrame(matched_rows)
    unmatched_api_df = pd.DataFrame(unmatched_api_rows)
    unmatched_db_df = pd.DataFrame(unmatched_db_rows)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        matched_df.to_excel(writer, sheet_name="Matched_Data", index=False)
        unmatched_api_df.to_excel(writer, sheet_name="Unmatched_API_Data", index=False)
        unmatched_db_df.to_excel(writer, sheet_name="Unmatched_Local_Data", index=False)

    print(f"Done | Matched: {len(matched_df)} | Unmatched API: {len(unmatched_api_df)} | Unmatched Local: {len(unmatched_db_df)}")

except FileNotFoundError:
    print(f"Error: {INPUT_FILE} not found.")
except Exception as e:
    print(f"Error occurred: {e}")