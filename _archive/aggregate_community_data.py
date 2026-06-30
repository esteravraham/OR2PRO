import pandas as pd

OFFICIAL_FILE = "Clean_Ramat_Gan_Pilot.xlsx"
FINAL_FILE = "Final_Ramat_Gan_Pilot.xlsx"
OUTPUT_FILE = "Community_Aggregation.xlsx"

official = pd.read_excel(OFFICIAL_FILE, sheet_name="Cleaned_Data")

try:
    final_data = pd.read_excel(FINAL_FILE, sheet_name="Final_Data")
except FileNotFoundError:
    final_data = official.copy()

community_fields = [
    "type",
    "ages",
    "activity_hours",
    "friday",
    "price",
    "community_notes",
    "external_link",
    "community_match_status",
    "community_phone",
    "phone_conflict"
]

for field in community_fields:
    if field not in final_data.columns:
        final_data[field] = "Unknown"

def clean_value(value):
    if pd.isna(value) or str(value).strip() in ["", "nan", "Unknown"]:
        return "Unknown"
    return str(value).strip()

def has_community_info(row):
    fields_to_check = [
        "type",
        "ages",
        "activity_hours",
        "friday",
        "price",
        "community_notes",
        "external_link"
    ]

    for field in fields_to_check:
        if clean_value(row.get(field, "Unknown")) != "Unknown":
            return True

    return False

aggregation_rows = []

for _, row in final_data.iterrows():
    has_info = has_community_info(row)

    aggregation_rows.append({
        # מידע רשמי
        "garden_id": row.get("id", "Unknown"),
        "name": row.get("name", "Unknown"),
        "city": row.get("city", "Unknown"),
        "address": row.get("address", "Unknown"),
        "official_phone": row.get("phone", "Unknown"),
        "ownership": row.get("ownership", "Unknown"),
        "sector": row.get("sector", "Unknown"),
        "manager": row.get("manager", "Unknown"),
        "license_status": row.get("license_status", "Unknown"),

        # מידע קהילתי קיים כרגע
        "has_community_info": has_info,
        "num_community_records": 1 if has_info else 0,
        "type_summary": clean_value(row.get("type", "Unknown")),
        "ages_summary": clean_value(row.get("ages", "Unknown")),
        "activity_hours_summary": clean_value(row.get("activity_hours", "Unknown")),
        "friday_summary": clean_value(row.get("friday", "Unknown")),
        "price_summary": clean_value(row.get("price", "Unknown")),
        "community_notes_summary": clean_value(row.get("community_notes", "Unknown")),
        "external_links_summary": clean_value(row.get("external_link", "Unknown")),

        # בקרת התאמה
        "community_match_status": clean_value(row.get("community_match_status", "Unknown")),
        "community_phone": clean_value(row.get("community_phone", "Unknown")),
        "phone_conflict": row.get("phone_conflict", False),

        # שדות עתידיים מהאתר
        "avg_rating": "Not available yet",
        "num_ratings": 0,
        "cameras_summary": "Not available yet",
        "protected_space_summary": "Not available yet",
        "staff_child_ratio_summary": "Not available yet",
        "parents_feeling_summary": "Not available yet",
        "reviews_summary": "Not available yet"
    })

aggregation_df = pd.DataFrame(aggregation_rows)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    aggregation_df.to_excel(writer, sheet_name="Community_Aggregation", index=False)
    final_data.to_excel(writer, sheet_name="Final_Data_Source", index=False)

print(
    f"✅ Done | Gardens: {len(aggregation_df)} | "
    f"With community info: {aggregation_df['has_community_info'].sum()} | "
    f"Without community info: {(~aggregation_df['has_community_info']).sum()}"
)