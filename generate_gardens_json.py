import pandas as pd
import json
import os
import shutil
from datetime import datetime

INPUT_FILE = "Parent_Input_Data_Structure.xlsx"
OUTPUT_FILE = "gardens.json"
BACKUP_FILE = "gardens_backup.json"

GARDEN_PROFILE_SHEET = "Garden_Profile"
PARENT_REVIEWS_SHEET = "Parent_Reviews_Template"
SUGGESTED_UPDATES_SHEET = "Suggested_Updates_Template"


def clean_value(value):
    if pd.isna(value):
        return None

    value = str(value).strip()

    if value in ["", "nan", "None", "Unknown", "Not available yet"]:
        return None

    return value


def clean_number(value):
    if pd.isna(value):
        return None

    value = str(value).replace(",", "").strip()

    if value in ["", "nan", "None", "Unknown"]:
        return None

    try:
        return float(value)
    except ValueError:
        return None


def is_approved(row):
    status = clean_value(row.get("status"))
    return status == "approved"


def summarize_reviews(reviews_df):
    if reviews_df.empty:
        return {
            "num_reviews": 0,
            "avg_rating": None,
            "price_min": None,
            "price_max": None,
            "price_avg": None,
            "price_without_friday": None,
            "price_with_friday": None,
            "price_details": [],
            "parent_feeling_summary": None,
            "review_texts": []
        }

    approved_reviews = reviews_df[reviews_df.apply(is_approved, axis=1)].copy()

    if approved_reviews.empty:
        return {
            "num_reviews": 0,
            "avg_rating": None,
            "price_min": None,
            "price_max": None,
            "price_avg": None,
            "price_without_friday": None,
            "price_with_friday": None,
            "price_details": [],
            "parent_feeling_summary": None,
            "review_texts": []
        }

    ratings = (
        approved_reviews["rating_1_to_5"].apply(clean_number).dropna()
        if "rating_1_to_5" in approved_reviews.columns
        else pd.Series(dtype=float)
    )

    prices = (
        approved_reviews["price_monthly"].apply(clean_number).dropna()
        if "price_monthly" in approved_reviews.columns
        else pd.Series(dtype=float)
    )

    prices_without_friday = (
        approved_reviews["price_without_friday"].apply(clean_number).dropna()
        if "price_without_friday" in approved_reviews.columns
        else pd.Series(dtype=float)
    )

    prices_with_friday = (
        approved_reviews["price_with_friday"].apply(clean_number).dropna()
        if "price_with_friday" in approved_reviews.columns
        else pd.Series(dtype=float)
    )

    price_details = (
        approved_reviews["price_details"]
        .apply(clean_value)
        .dropna()
        .tolist()
        if "price_details" in approved_reviews.columns
        else []
    )

    feelings = (
        approved_reviews["parent_feeling"]
        .apply(clean_value)
        .dropna()
        .value_counts()
        .to_dict()
        if "parent_feeling" in approved_reviews.columns
        else {}
    )

    review_texts = (
        approved_reviews["review_text"]
        .apply(clean_value)
        .dropna()
        .tolist()
        if "review_text" in approved_reviews.columns
        else []
    )

    return {
        "num_reviews": int(len(approved_reviews)),
        "avg_rating": round(float(ratings.mean()), 2) if not ratings.empty else None,

        "price_min": int(prices.min()) if not prices.empty else None,
        "price_max": int(prices.max()) if not prices.empty else None,
        "price_avg": int(prices.mean()) if not prices.empty else None,

        "price_without_friday": int(prices_without_friday.mean()) if not prices_without_friday.empty else None,
        "price_with_friday": int(prices_with_friday.mean()) if not prices_with_friday.empty else None,
        "price_details": price_details,

        "parent_feeling_summary": feelings if feelings else None,
        "review_texts": review_texts
    }


def summarize_suggestions(suggestions_df):
    if suggestions_df.empty:
        return {
            "approved_updates": {},
            "pending_suggestions_count": 0
        }

    approved_suggestions = suggestions_df[suggestions_df.apply(is_approved, axis=1)].copy()

    pending_suggestions = (
        suggestions_df[suggestions_df["status"].apply(clean_value) == "pending_review"].copy()
        if "status" in suggestions_df.columns
        else pd.DataFrame()
    )

    approved_by_field = {}

    for _, row in approved_suggestions.iterrows():
        field_name = clean_value(row.get("field_name"))
        suggested_value = clean_value(row.get("suggested_value"))

        if field_name and suggested_value:
            approved_by_field[field_name] = suggested_value

    return {
        "approved_updates": approved_by_field,
        "pending_suggestions_count": int(len(pending_suggestions))
    }


def build_search_text(values):
    clean_values = [clean_value(value) for value in values]
    clean_values = [value for value in clean_values if value]
    return " ".join(clean_values).lower()


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ {INPUT_FILE} not found")
        return

    garden_profile = pd.read_excel(INPUT_FILE, sheet_name=GARDEN_PROFILE_SHEET)

    try:
        parent_reviews = pd.read_excel(INPUT_FILE, sheet_name=PARENT_REVIEWS_SHEET)
    except Exception:
        parent_reviews = pd.DataFrame()

    try:
        suggested_updates = pd.read_excel(INPUT_FILE, sheet_name=SUGGESTED_UPDATES_SHEET)
    except Exception:
        suggested_updates = pd.DataFrame()

    if os.path.exists(OUTPUT_FILE):
        shutil.copyfile(OUTPUT_FILE, BACKUP_FILE)

    gardens = []

    profile_fields = [
        "garden_type",
        "ages",
        "activity_hours",
        "friday",
        "external_link",

        "education_type",
        "education_type_other",
        "religious_orientation",
        "religious_orientation_other",
        "nutrition_type",
        "nutrition_type_other",

        "has_cameras",
        "cameras_open_to_parents",
        "has_protected_space",

        "profile_data_status",
        "last_verified_at",
        "verified_by"
    ]

    for _, garden in garden_profile.iterrows():
        garden_id = str(garden.get("garden_id"))

        garden_reviews = (
            parent_reviews[parent_reviews["garden_id"].astype(str) == garden_id]
            if not parent_reviews.empty and "garden_id" in parent_reviews.columns
            else pd.DataFrame()
        )

        garden_suggestions = (
            suggested_updates[suggested_updates["garden_id"].astype(str) == garden_id]
            if not suggested_updates.empty and "garden_id" in suggested_updates.columns
            else pd.DataFrame()
        )

        review_summary = summarize_reviews(garden_reviews)
        suggestion_summary = summarize_suggestions(garden_suggestions)
        approved_updates = suggestion_summary.get("approved_updates", {})

        official_name = clean_value(garden.get("name"))
        official_phone = clean_value(garden.get("official_phone"))

        display_name = approved_updates.get("display_name") or official_name
        display_phone = approved_updates.get("display_phone") or official_phone

        profile = {}

        for field in profile_fields:
            value = clean_value(garden.get(field))

            # אם יש הצעה מאושרת לאותו שדה, היא גוברת על הערך מהפרופיל
            if field in approved_updates:
                value = approved_updates[field]

            profile[field] = value

        has_profile_community_info = any(
            profile.get(field) is not None
            for field in [
                "garden_type",
                "ages",
                "activity_hours",
                "friday",
                "external_link",
                "education_type",
                "religious_orientation",
                "nutrition_type",
                "has_cameras",
                "cameras_open_to_parents",
                "has_protected_space"
            ]
        )

        has_community_info = (
            has_profile_community_info
            or review_summary["num_reviews"] > 0
            or len(approved_updates) > 0
            or suggestion_summary["pending_suggestions_count"] > 0
        )

        search_text = build_search_text([
            display_name,
            official_name,
            garden.get("city"),
            garden.get("address"),
            garden.get("manager"),
            garden.get("ownership"),
            garden.get("sector"),
            profile.get("garden_type"),
            profile.get("ages"),
            profile.get("education_type"),
            profile.get("religious_orientation"),
            profile.get("nutrition_type")
        ])

        garden_json = {
            "id": garden_id,

            "name": display_name,
            "display_name": display_name,
            "official_name": official_name,

            "city": clean_value(garden.get("city")),
            "address": clean_value(garden.get("address")),

            "phone": display_phone,
            "display_phone": display_phone,
            "official_phone": official_phone,

            "ownership": clean_value(garden.get("ownership")),
            "sector": clean_value(garden.get("sector")),
            "manager": clean_value(garden.get("manager")),
            "license_status": clean_value(garden.get("license_status")),

            "profile": profile,

            "community": {
                "has_community_info": has_community_info,
                "reviews": review_summary,
                "suggestions": suggestion_summary
            },

            "search_text": search_text
        }

        gardens.append(garden_json)

    output = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": INPUT_FILE,
        "total_gardens": len(gardens),
        "gardens": gardens
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"✅ Done | Created {OUTPUT_FILE} with {len(gardens)} gardens")

    if os.path.exists(BACKUP_FILE):
        print(f"ℹ️ Previous gardens.json was backed up as {BACKUP_FILE}")


if __name__ == "__main__":
    main()
    