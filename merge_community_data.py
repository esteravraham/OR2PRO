import pandas as pd
import re
from difflib import SequenceMatcher

OFFICIAL_FILE = "Clean_Ramat_Gan_Pilot.xlsx"
COMMUNITY_FILE = "Clean_Community_Data.xlsx"
OUTPUT_FILE = "Final_Ramat_Gan_Pilot.xlsx"

official = pd.read_excel(OFFICIAL_FILE, sheet_name="Cleaned_Data")
community = pd.read_excel(COMMUNITY_FILE, sheet_name="Clean_Community_Data")

def normalize_text(value, remove_numbers=True):
    if pd.isna(value) or str(value).strip() == "Unknown":
        return ""

    value = str(value).strip().lower()

    if remove_numbers:
        value = re.sub(r"\d+", "", value)

    value = re.sub(r'[-–—״"\'.,()]', " ", value)
    value = re.sub(r"\s+", " ", value)

    return value.strip()

def text_similarity(value1, value2, remove_numbers=True):
    value1 = normalize_text(value1, remove_numbers=remove_numbers)
    value2 = normalize_text(value2, remove_numbers=remove_numbers)

    if value1 == "" or value2 == "":
        return 0

    if value1 == value2:
        return 1

    if min(len(value1), len(value2)) >= 4:
        if value1 in value2 or value2 in value1:
            return 0.9

    return SequenceMatcher(None, value1, value2).ratio()

def phone_equal(phone1, phone2):
    if pd.isna(phone1) or pd.isna(phone2):
        return False

    phone1 = re.sub(r"[^0-9]", "", str(phone1))
    phone2 = re.sub(r"[^0-9]", "", str(phone2))

    if phone1 == "" or phone2 == "":
        return False

    return phone1 == phone2

def match_score(official_row, community_row):
    score = 0
    reasons = []

    address_sim = text_similarity(
    official_row.get("address", ""),
    community_row.get("address", ""),
    remove_numbers=False
)
    name_sim = text_similarity(
        official_row.get("name", ""),
        community_row.get("name", "")
    )

    manager_sim = text_similarity(
        official_row.get("manager", ""),
        community_row.get("manager", "")
    )

    if address_sim >= 0.9:
        score += 3
        reasons.append("address_match")
    elif address_sim >= 0.75:
        score += 2
        reasons.append("address_similar")

    if name_sim >= 0.85:
        score += 3
        reasons.append("name_match")
    elif name_sim >= 0.7:
        score += 2
        reasons.append("name_similar")

    if phone_equal(official_row.get("phone", ""), community_row.get("phone", "")):
        score += 4
        reasons.append("phone_match")

    if manager_sim >= 0.85:
        score += 2
        reasons.append("manager_match")

    return score, reasons

def is_phone_conflict(official_phone, community_phone):
    if pd.isna(community_phone) or str(community_phone).strip() in ["", "Unknown"]:
        return False

    if pd.isna(official_phone) or str(official_phone).strip() in ["", "Unknown"]:
        return False

    return not phone_equal(official_phone, community_phone)

community_fields = [
    "type",
    "ages",
    "activity_hours",
    "friday",
    "price",
    "community_notes",
    "external_link"
]

for col in community_fields:
    official[col] = ""

official["community_match_status"] = ""
official["community_phone"] = ""
official["phone_conflict"] = False
official["match_score"] = 0
official["match_reasons"] = ""

possible_matches = []
not_matched = []
matched_count = 0

for _, community_row in community.iterrows():
    candidates = []

    for _, official_row in official.iterrows():
        score, reasons = match_score(official_row, community_row)

        if score >= 3:
            candidates.append({
                "official_index": official_row.name,
                "official_id": official_row.get("id", ""),
                "official_name": official_row.get("name", ""),
                "official_address": official_row.get("address", ""),
                "official_phone": official_row.get("phone", ""),
                "community_name": community_row.get("name", ""),
                "community_address": community_row.get("address", ""),
                "community_phone": community_row.get("phone", ""),
                "score": score,
                "reasons": " | ".join(reasons)
            })

    candidates = sorted(candidates, key=lambda x: x["score"], reverse=True)

    if candidates and candidates[0]["score"] >= 6:
        best = candidates[0]
        match_index = best["official_index"]

        for field in community_fields:
            official.loc[match_index, field] = community_row.get(field, "")

        official.loc[match_index, "community_match_status"] = "matched_by_score"
        official.loc[match_index, "community_phone"] = community_row.get("phone", "")
        official.loc[match_index, "phone_conflict"] = is_phone_conflict(
            official.loc[match_index, "phone"],
            community_row.get("phone", "")
        )
        official.loc[match_index, "match_score"] = best["score"]
        official.loc[match_index, "match_reasons"] = best["reasons"]

        matched_count += 1

    elif candidates:
        for candidate in candidates[:5]:
            possible_matches.append(candidate)

    else:
        not_matched.append(community_row)

possible_matches_df = pd.DataFrame(possible_matches)
not_matched_df = pd.DataFrame(not_matched)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    official.to_excel(writer, sheet_name="Final_Data", index=False)
    possible_matches_df.to_excel(writer, sheet_name="Possible_Matches", index=False)
    not_matched_df.to_excel(writer, sheet_name="Community_Not_Matched", index=False)

print(
    f"✅ Done | Final rows: {len(official)} | "
    f"Matched: {matched_count} | "
    f"Possible: {len(possible_matches_df)} | "
    f"Not matched: {len(not_matched_df)}"
)