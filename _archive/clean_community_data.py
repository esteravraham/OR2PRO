import pandas as pd
import re

SOURCE_FILE = "Ramat_Gan_Pilot.xlsx"
OUTPUT_FILE = "Clean_Community_Data.xlsx"

community = pd.read_excel(SOURCE_FILE, sheet_name="Community_Data")

community = community.rename(columns={
    "שם הגן": "name",
    "כתובת": "address",
    "מנהל/ת": "manager",
    "טלפון": "phone",
    "משפחתון/גן": "type",
    "גילאים": "ages",
    "שעות פעילות": "activity_hours",
    "שישי": "friday",
    "מחיר": "price",
    "הערות נוספות - המלצה/דיס המלצה": "community_notes",
    "קישור לדף הגן": "external_link"
})

def clean_text(value):
    if pd.isna(value):
        return "Unknown"
    value = str(value).strip()
    value = re.sub(r"\s+", " ", value)
    return value if value else "Unknown"

def normalize_text(value, remove_numbers=True):
    if pd.isna(value) or value == "Unknown":
        return ""
    value = str(value).strip()
    if remove_numbers:
        value = re.sub(r"\d+", "", value)
    value = re.sub(r'[-–—״"\'.,()]', " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()

def clean_phone(value):
    if pd.isna(value):
        return "Unknown"

    value = str(value).strip()

    # טיפול במקרה שאקסל הפך טלפון למספר מדעי
    try:
        if "E+" in value.upper():
            value = str(int(float(value)))
    except:
        pass

    value = re.sub(r"[^0-9]", "", value)

    if value == "":
        return "Unknown"

    if value.startswith("972"):
        value = "0" + value[3:]

    if not value.startswith("0"):
        value = "0" + value

    return value

for col in community.columns:
    if col != "phone":
        community[col] = community[col].apply(clean_text)

community["phone"] = community["phone"].apply(clean_phone)

community["name_norm"] = community["name"].apply(lambda x: normalize_text(x, remove_numbers=True))
community["address_norm"] = community["address"].apply(lambda x: normalize_text(x, remove_numbers=False))
community["manager_norm"] = community["manager"].apply(lambda x: normalize_text(x, remove_numbers=True))

community = community.drop_duplicates()

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    community.to_excel(writer, sheet_name="Clean_Community_Data", index=False)

print(f"✅ Done | Community rows: {len(community)}")