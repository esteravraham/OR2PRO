import json
import os
import math
import numbers
import pandas as pd
from collections import deque


GARDENS_JSON_FILE = "gardens.json"
PARENTS_INPUT_FILE = "Parents_Input.xlsx"
GARDENS_INPUT_FILE = "Gardens_Matching_Input.xlsx"
OUTPUT_FILE = "Gale_Shapley_Matching_Result.xlsx"
FULL_OUTPUT_FILE = "Gale_Shapley_Matching_Result_FULL.xlsx"

DEFAULT_CAPACITY = 1
DEFAULT_MAX_DISTANCE_KM = 2.5
DEFAULT_PREFERRED_DISTANCE_KM = 0.5
MAX_CANDIDATE_GARDENS_PER_PARENT = 30
MIN_PARENT_SCORE_FOR_PREFERENCE = 0

DEBUG_OUTPUT = True

TZAHARON_HARD_IMPORTANCE_THRESHOLD = 4
GENDER_HARD_IMPORTANCE_THRESHOLD = 4
RELIGIOUS_HARD_IMPORTANCE_THRESHOLD = 5
RELIGIOUS_SUBSTREAM_HARD_IMPORTANCE_THRESHOLD = 5
SOCIAL_FRIEND_BONUS_PER_IMPORTANCE = 12
SOCIAL_FRIEND_MAX_BONUS = 60

GARDEN_COLUMNS = [
    "garden_id", "garden_name", "address", "neighborhood", "x", "y", "capacity",
    "min_age_months", "max_age_months", "sector", "education_type",
    "religious_orientation", "religious_substream", "gender_composition",
    "activity_language", "price_avg", "friday", "has_tzaharon", "tzaharon_until",
    "has_protected_space"
]

REQUIRED_PARENT_COLUMNS = [
    "parent_id", "parent_name", "child_name", "child_age_months", "home_address",
    "home_neighborhood", "home_lat", "home_lon", "max_distance_km", "max_price",
    "preferred_activity_language", "preferred_education_type",
    "preferred_religious_orientation", "preferred_sector", "needs_friday",
    "needs_protected_space", "declared_sibling_in_garden", "sibling_garden_id",
    "sibling_verification_status", "preferred_garden_1", "preferred_garden_2",
    "preferred_garden_3", "importance_distance", "importance_price",
    "importance_activity_language", "importance_education_type",
    "importance_religious_orientation", "importance_sector", "importance_friday",
    "importance_protected_space"
]

OPTIONAL_PARENT_COLUMNS = [
    "allow_manual_far_preference",
    "max_manual_exception_distance_km",
    "preferred_distance_km",
    "distance_choice_label",
    "needs_tzaharon",
    "importance_tzaharon",
    "preferred_religious_substream",
    "importance_religious_substream",
    "preferred_gender_composition",
    "importance_gender_composition",
    "friend_request_1",
    "friend_request_2",
    "importance_same_friend",
    "willing_to_trade_distance_for_friend",
    "child_birth_month"
]

PROPOSAL_LOG_COLUMNS = [
    "iteration", "event_type", "parent_id", "proposal_number_for_parent",
    "garden_id", "decision", "reason", "capacity", "matches_before",
    "matches_after", "rejected_in_this_step"
]


def clean_value(value):
    if pd.isna(value):
        return ""

    if isinstance(value, numbers.Number) and not isinstance(value, bool):
        try:
            number = float(value)
            if number.is_integer():
                return str(int(number))
        except (TypeError, ValueError):
            pass

    text = str(value).strip()

    if text.lower() in ["", "nan", "none", "null"]:
        return ""

    if text.endswith(".0"):
        possible_number = text[:-2]
        if possible_number.isdigit():
            return possible_number

    return text


def normalize_text(value):
    return clean_value(value).lower().strip()


def first_non_empty(*values):
    for value in values:
        cleaned = clean_value(value)
        if cleaned != "":
            return cleaned
    return ""


def to_number(value, default=None):
    text = clean_value(value)

    if text == "":
        return default

    text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return default


def to_int(value, default=0):
    number = to_number(value, default=None)

    if number is None:
        return default

    return int(number)


def is_yes(value):
    text = normalize_text(value)
    return text in ["כן", "yes", "true", "1", "y", "יש", "קיים", "approved"] or "כן" in text


def is_no_value(value):
    text = normalize_text(value)
    return text in ["לא", "no", "false", "0", "אין", "ללא"]


def contains_match(preferred_value, actual_value):
    preferred = normalize_text(preferred_value)
    actual = normalize_text(actual_value)

    if preferred == "" or actual == "":
        return False

    return preferred in actual or actual in preferred


def is_any_value(value):
    text = normalize_text(value)
    return text in ["", "לא משנה", "any", "all", "גמיש", "אין העדפה", "לא חשוב"]


def hard_preference_match(preferred_value, actual_value):
    if is_any_value(preferred_value):
        return True
    return contains_match(preferred_value, actual_value)


def hard_constraint_check(parent_row, garden_row, check_distance=True):
    compatible, age_reason = age_is_compatible(parent_row, garden_row)
    if not compatible:
        return False, age_reason

    if check_distance:
        distance_info = get_distance_info(parent_row, garden_row)
        if not distance_info["within_normal_distance"]:
            return False, distance_info["distance_reason"]

    needs_tzaharon = is_yes(parent_row.get("needs_tzaharon", ""))
    importance_tzaharon = get_importance(parent_row, "importance_tzaharon", default=1)
    if needs_tzaharon and importance_tzaharon >= TZAHARON_HARD_IMPORTANCE_THRESHOLD:
        if not is_yes(garden_row.get("has_tzaharon", "")):
            return False, "tzaharon_required_but_missing"

    preferred_gender = parent_row.get("preferred_gender_composition", "")
    importance_gender = get_importance(parent_row, "importance_gender_composition", default=1)
    if not is_any_value(preferred_gender) and importance_gender >= GENDER_HARD_IMPORTANCE_THRESHOLD:
        if not hard_preference_match(preferred_gender, garden_row.get("gender_composition", "")):
            return False, "gender_composition_required_but_not_matching"

    preferred_substream = parent_row.get("preferred_religious_substream", "")
    importance_substream = get_importance(parent_row, "importance_religious_substream", default=1)
    if not is_any_value(preferred_substream) and importance_substream >= RELIGIOUS_SUBSTREAM_HARD_IMPORTANCE_THRESHOLD:
        if not hard_preference_match(preferred_substream, garden_row.get("religious_substream", "")):
            return False, "religious_substream_required_but_not_matching"

    return True, "hard_constraints_passed"


def get_parent_requested_friend_ids(parent_row):
    return [
        clean_value(parent_row.get("friend_request_1", "")),
        clean_value(parent_row.get("friend_request_2", ""))
    ]


def build_mutual_friend_pairs(parents_df):
    requests = {}
    for _, row in parents_df.iterrows():
        parent_id = clean_value(row.get("parent_id", ""))
        if not parent_id:
            continue
        requests[parent_id] = set(friend_id for friend_id in get_parent_requested_friend_ids(row) if friend_id)

    mutual_pairs = set()
    for parent_id, requested_ids in requests.items():
        for friend_id in requested_ids:
            if friend_id in requests and parent_id in requests.get(friend_id, set()):
                mutual_pairs.add(tuple(sorted([parent_id, friend_id])))

    return sorted(mutual_pairs)


def load_gardens_json():
    if not os.path.exists(GARDENS_JSON_FILE):
        return []

    with open(GARDENS_JSON_FILE, "r", encoding="utf-8") as file:
        data = json.load(file)

    if isinstance(data, dict):
        return data.get("gardens", [])

    if isinstance(data, list):
        return data

    return []


def build_gardens_dataframe_from_json():
    gardens = load_gardens_json()
    rows = []

    for garden in gardens:
        profile = garden.get("profile", {}) or {}
        community = garden.get("community", {}) or {}
        reviews = community.get("reviews", {}) or {}

        rows.append({
            "garden_id": first_non_empty(garden.get("garden_id"), garden.get("id"), garden.get("symbol")),
            "garden_name": first_non_empty(garden.get("garden_name"), garden.get("display_name"), garden.get("name"), garden.get("official_name")),
            "address": first_non_empty(garden.get("address"), profile.get("address")),
            "neighborhood": first_non_empty(garden.get("neighborhood"), profile.get("neighborhood")),
            "x": first_non_empty(garden.get("x"), garden.get("longitude"), profile.get("x"), profile.get("longitude")),
            "y": first_non_empty(garden.get("y"), garden.get("latitude"), profile.get("y"), profile.get("latitude")),
            "capacity": to_int(garden.get("capacity"), default=DEFAULT_CAPACITY),
            "min_age_months": first_non_empty(profile.get("min_age_months"), garden.get("min_age_months")),
            "max_age_months": first_non_empty(profile.get("max_age_months"), garden.get("max_age_months")),
            "sector": first_non_empty(garden.get("sector"), profile.get("sector")),
            "education_type": first_non_empty(garden.get("education_type"), profile.get("education_type"), profile.get("garden_type")),
            "religious_orientation": first_non_empty(garden.get("religious_orientation"), profile.get("religious_orientation")),
            "religious_substream": first_non_empty(garden.get("religious_substream"), profile.get("religious_substream")),
            "gender_composition": first_non_empty(garden.get("gender_composition"), profile.get("gender_composition")),
            "activity_language": first_non_empty(garden.get("activity_language"), profile.get("activity_language")),
            "price_avg": first_non_empty(garden.get("price_avg"), profile.get("price_avg"), reviews.get("price_avg")),
            "friday": first_non_empty(garden.get("friday"), profile.get("friday")),
            "has_tzaharon": first_non_empty(garden.get("has_tzaharon"), profile.get("has_tzaharon")),
            "tzaharon_until": first_non_empty(garden.get("tzaharon_until"), profile.get("tzaharon_until")),
            "has_protected_space": first_non_empty(garden.get("has_protected_space"), profile.get("has_protected_space"))
        })

    return pd.DataFrame(rows, columns=GARDEN_COLUMNS)


def ensure_gardens_input_file():
    gardens_df = build_gardens_dataframe_from_json()
    if gardens_df.empty:
        raise FileNotFoundError(f"Missing {GARDENS_JSON_FILE}")
    return gardens_df


def validate_parents_input(parents_df):
    missing_columns = [
        column for column in REQUIRED_PARENT_COLUMNS
        if column not in parents_df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Parents_Input.xlsx is missing these columns: "
            + ", ".join(missing_columns)
        )

    for column in OPTIONAL_PARENT_COLUMNS:
        if column not in parents_df.columns:
            parents_df[column] = ""


def haversine_km(lat1, lon1, lat2, lon2):
    lat1 = to_number(lat1)
    lon1 = to_number(lon1)
    lat2 = to_number(lat2)
    lon2 = to_number(lon2)

    if None in [lat1, lon1, lat2, lon2]:
        return None

    radius = 6371
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)

    a = (
        math.sin(delta_phi / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    )

    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    return radius * c



def get_distance_info(parent_row, garden_row):
    distance = haversine_km(
        parent_row.get("home_lat", ""),
        parent_row.get("home_lon", ""),
        garden_row.get("y", ""),
        garden_row.get("x", "")
    )

    max_distance = to_number(parent_row.get("max_distance_km"), default=DEFAULT_MAX_DISTANCE_KM)
    preferred_distance = to_number(parent_row.get("preferred_distance_km"), default=DEFAULT_PREFERRED_DISTANCE_KM)

    parent_neighborhood = normalize_text(parent_row.get("home_neighborhood", ""))
    garden_neighborhood = normalize_text(garden_row.get("neighborhood", ""))

    same_neighborhood = (
        parent_neighborhood != ""
        and garden_neighborhood != ""
        and parent_neighborhood == garden_neighborhood
    )

    if distance is not None:
        within_normal = distance <= max_distance
        within_preferred = distance <= preferred_distance
        method = "coordinates"
        reason = "within_normal_distance" if within_normal else "above_normal_distance"
    elif same_neighborhood:
        within_normal = True
        within_preferred = False
        method = "same_neighborhood_fallback"
        reason = "same_neighborhood_no_coordinates"
    else:
        within_normal = False
        within_preferred = False
        method = "distance_missing"
        reason = "distance_missing_and_neighborhood_not_equal"

    return {
        "distance_km": distance,
        "max_distance_km": max_distance,
        "preferred_distance_km": preferred_distance,
        "within_normal_distance": within_normal,
        "within_preferred_distance": within_preferred,
        "same_neighborhood": same_neighborhood,
        "distance_method": method,
        "distance_reason": reason
    }

def age_is_compatible(parent_row, garden_row):
    child_age = to_number(parent_row.get("child_age_months"), default=None)
    min_age = to_number(garden_row.get("min_age_months"), default=None)
    max_age = to_number(garden_row.get("max_age_months"), default=None)

    if child_age is None:
        return True, "age_missing"

    if min_age is None and max_age is None:
        return True, "age_range_missing"

    if min_age is not None and child_age < min_age:
        return False, "child_too_young"

    if max_age is not None and child_age > max_age:
        return False, "child_too_old"

    return True, "age_match"


def get_importance(row, column_name, default=3):
    value = to_int(row.get(column_name, default), default=default)

    if value < 1:
        return 1

    if value > 5:
        return 5

    return value


def get_manual_preference_ids(parent_row):
    return [
        clean_value(parent_row.get("preferred_garden_1", "")),
        clean_value(parent_row.get("preferred_garden_2", "")),
        clean_value(parent_row.get("preferred_garden_3", ""))
    ]


def get_manual_preference_rank(parent_row, garden_id):
    garden_id = clean_value(garden_id)

    for index, preferred_garden_id in enumerate(get_manual_preference_ids(parent_row), start=1):
        if preferred_garden_id and preferred_garden_id == garden_id:
            return index

    return None


def can_use_manual_exception(parent_row, garden_row):
    if not is_yes(parent_row.get("allow_manual_far_preference", "")):
        return False, "parent_did_not_allow_manual_far_preference"

    garden_id = clean_value(garden_row.get("garden_id", ""))
    manual_rank = get_manual_preference_rank(parent_row, garden_id)

    if manual_rank is None:
        return False, "garden_not_in_manual_preferences"

    distance_info = get_distance_info(parent_row, garden_row)
    distance = distance_info["distance_km"]

    if distance is None:
        return False, "distance_missing_for_manual_exception"

    max_exception_distance = to_number(
        parent_row.get("max_manual_exception_distance_km"),
        default=None
    )

    if max_exception_distance is None:
        return False, "manual_exception_max_distance_missing"

    if distance > max_exception_distance:
        return False, "above_manual_exception_distance"

    compatible, age_reason = age_is_compatible(parent_row, garden_row)

    if not compatible:
        return False, age_reason

    return True, "manual_exception_allowed"



def calculate_parent_fit_score(parent_row, garden_row, skip_distance=False):
    score = 0.0
    reasons = []

    hard_ok, hard_reason = hard_constraint_check(parent_row, garden_row, check_distance=not skip_distance)
    if not hard_ok:
        return 0.0, [hard_reason], False

    compatible, age_reason = age_is_compatible(parent_row, garden_row)
    reasons.append(age_reason)

    distance_info = get_distance_info(parent_row, garden_row)
    manual_rank = get_manual_preference_rank(parent_row, garden_row.get("garden_id"))

    if manual_rank is not None:
        score += {1: 300, 2: 250, 3: 200}.get(manual_rank, 0)
        reasons.append(f"manual_preference_{manual_rank}")

    importance_distance = get_importance(parent_row, "importance_distance")
    distance = distance_info["distance_km"]
    max_distance = distance_info["max_distance_km"]
    preferred_distance = distance_info["preferred_distance_km"]

    if distance is not None and max_distance is not None and max_distance > 0:
        distance_score = (1 - distance / max_distance) * importance_distance * 20
        score += max(distance_score, 0)
        reasons.append(f"distance_match_{distance:.2f}_km")

        if preferred_distance is not None and distance <= preferred_distance:
            preferred_bonus = importance_distance * 25
            score += preferred_bonus
            reasons.append(f"preferred_distance_bonus_up_to_{preferred_distance:.2f}_km")
    elif distance_info["same_neighborhood"]:
        score += importance_distance * 20
        reasons.append("neighborhood_match")

    importance_price = get_importance(parent_row, "importance_price")
    max_price = to_number(parent_row.get("max_price"), default=None)
    garden_price = to_number(garden_row.get("price_avg"), default=None)

    if max_price is not None and garden_price is not None:
        if garden_price <= max_price:
            score += importance_price * 15
            reasons.append("price_match")
        else:
            reasons.append("price_above_budget")

    if contains_match(parent_row.get("preferred_activity_language"), garden_row.get("activity_language")):
        score += get_importance(parent_row, "importance_activity_language") * 15
        reasons.append("activity_language_match")

    if contains_match(parent_row.get("preferred_education_type"), garden_row.get("education_type")):
        score += get_importance(parent_row, "importance_education_type") * 15
        reasons.append("education_type_match")

    if contains_match(parent_row.get("preferred_religious_orientation"), garden_row.get("religious_orientation")):
        score += get_importance(parent_row, "importance_religious_orientation") * 15
        reasons.append("religious_orientation_match")

    if contains_match(parent_row.get("preferred_religious_substream"), garden_row.get("religious_substream")):
        score += get_importance(parent_row, "importance_religious_substream") * 18
        reasons.append("religious_substream_match")

    if contains_match(parent_row.get("preferred_gender_composition"), garden_row.get("gender_composition")):
        score += get_importance(parent_row, "importance_gender_composition") * 18
        reasons.append("gender_composition_match")

    if contains_match(parent_row.get("preferred_sector"), garden_row.get("sector")):
        score += get_importance(parent_row, "importance_sector") * 10
        reasons.append("sector_match")

    if is_yes(parent_row.get("needs_tzaharon")):
        if is_yes(garden_row.get("has_tzaharon")):
            score += get_importance(parent_row, "importance_tzaharon") * 20
            reasons.append("tzaharon_match")
        else:
            reasons.append("tzaharon_missing_soft")

    if is_yes(parent_row.get("needs_friday")):
        if clean_value(garden_row.get("friday")) and not is_no_value(garden_row.get("friday")):
            score += get_importance(parent_row, "importance_friday") * 10
            reasons.append("friday_match")
        else:
            reasons.append("friday_missing")

    if is_yes(parent_row.get("needs_protected_space")):
        if is_yes(garden_row.get("has_protected_space")):
            score += get_importance(parent_row, "importance_protected_space") * 10
            reasons.append("protected_space_match")
        else:
            reasons.append("protected_space_missing")

    return score, reasons, True


def calculate_garden_priority_score(parent_row, garden_row, parent_fit_score):
    import random
    score = 0.0
    reasons = []

    hard_ok, hard_reason = hard_constraint_check(parent_row, garden_row)
    if not hard_ok:
        return -1.0, [hard_reason]

    garden_id = clean_value(garden_row.get("garden_id", ""))

    # 1. גיל תואם
    compatible, age_reason = age_is_compatible(parent_row, garden_row)
    if compatible:
        score += 3
        reasons.append("age_match")

    # 2. אח/אחות באותו גן
    sibling_garden_id = clean_value(parent_row.get("sibling_garden_id", ""))
    has_sibling = is_yes(parent_row.get("declared_sibling_in_garden", ""))
    if has_sibling and sibling_garden_id == garden_id:
        score += 2
        reasons.append("sibling_in_garden")

    # 3. אותה שכונה
    distance_info = get_distance_info(parent_row, garden_row)
    if distance_info["same_neighborhood"]:
        score += 1
        reasons.append("same_neighborhood")

    # tiebreaker: הגרלה
    score += random.uniform(0, 0.9)
    reasons.append("random_tiebreaker")

    return score, reasons



def get_candidate_gardens_for_parent(parent_row, gardens_records, all_garden_ids=None):
    parent_id = clean_value(parent_row.get("parent_id", ""))
    manual_ids = [garden_id for garden_id in get_manual_preference_ids(parent_row) if garden_id]

    total_gardens = len(gardens_records)
    active_gardens = []
    age_compatible_active_gardens = []
    within_normal_distance_gardens = []
    hard_constraint_passed_gardens = []
    candidate_rows = []
    manual_included = []
    manual_excluded_due_distance = []
    manual_excluded_due_age = []
    manual_excluded_due_other_hard_constraint = []
    manual_not_found = []
    hard_rejection_counts = {}

    if all_garden_ids is None:
        all_garden_ids = set(clean_value(row.get("garden_id", "")) for row in gardens_records)

    for manual_id in manual_ids:
        if manual_id not in all_garden_ids:
            manual_not_found.append(manual_id)

    for garden_row in gardens_records:
        garden_id = clean_value(garden_row.get("garden_id", ""))
        capacity = to_int(garden_row.get("capacity", DEFAULT_CAPACITY), default=DEFAULT_CAPACITY)

        if not garden_id or capacity <= 0:
            continue

        active_gardens.append(garden_id)

        compatible, age_reason = age_is_compatible(parent_row, garden_row)

        if not compatible:
            hard_rejection_counts[age_reason] = hard_rejection_counts.get(age_reason, 0) + 1
            if garden_id in manual_ids:
                manual_excluded_due_age.append(garden_id)
            continue

        age_compatible_active_gardens.append(garden_id)

        distance_info = get_distance_info(parent_row, garden_row)
        if distance_info["within_normal_distance"]:
            within_normal_distance_gardens.append(garden_id)
        else:
            hard_rejection_counts[distance_info["distance_reason"]] = hard_rejection_counts.get(distance_info["distance_reason"], 0) + 1
            if garden_id in manual_ids:
                manual_excluded_due_distance.append(garden_id)
            continue

        hard_ok, hard_reason = hard_constraint_check(parent_row, garden_row)
        if not hard_ok:
            hard_rejection_counts[hard_reason] = hard_rejection_counts.get(hard_reason, 0) + 1
            if garden_id in manual_ids:
                manual_excluded_due_other_hard_constraint.append(f"{garden_id}:{hard_reason}")
            continue

        hard_constraint_passed_gardens.append(garden_id)
        candidate_rows.append(garden_row)
        if garden_id in manual_ids:
            manual_included.append(garden_id)

    # ── גיבוי שלבי: אם יש פחות מ-3 מועמדים, מרחיבים חיפוש ──────────────────
    MIN_CANDIDATES = 10
    fallback_stage = ""

    if len(candidate_rows) < MIN_CANDIDATES:
        # שלב א: כל העיר, שומרים גיל + צהרון קשיח
        existing_ids = {clean_value(r.get("garden_id", "")) for r in candidate_rows}
        needs_tz = is_yes(parent_row.get("needs_tzaharon", ""))
        imp_tz = get_importance(parent_row, "importance_tzaharon", default=1)
        for garden_row in gardens_records:
            garden_id = clean_value(garden_row.get("garden_id", ""))
            if not garden_id or garden_id in existing_ids:
                continue
            if to_int(garden_row.get("capacity", DEFAULT_CAPACITY), default=DEFAULT_CAPACITY) <= 0:
                continue
            compatible, _ = age_is_compatible(parent_row, garden_row)
            if not compatible:
                continue
            if needs_tz and imp_tz >= TZAHARON_HARD_IMPORTANCE_THRESHOLD:
                if not is_yes(garden_row.get("has_tzaharon", "")):
                    continue
            candidate_rows.append(garden_row)
            existing_ids.add(garden_id)
        fallback_stage = "city_wide_age_tzaharon"

    if len(candidate_rows) < MIN_CANDIDATES:
        # שלב ב: כל העיר, גיל בלבד
        existing_ids = {clean_value(r.get("garden_id", "")) for r in candidate_rows}
        for garden_row in gardens_records:
            garden_id = clean_value(garden_row.get("garden_id", ""))
            if not garden_id or garden_id in existing_ids:
                continue
            if to_int(garden_row.get("capacity", DEFAULT_CAPACITY), default=DEFAULT_CAPACITY) <= 0:
                continue
            compatible, _ = age_is_compatible(parent_row, garden_row)
            if not compatible:
                continue
            candidate_rows.append(garden_row)
            existing_ids.add(garden_id)
        fallback_stage = "city_wide_age_only"

    if len(candidate_rows) == 0:
        # שלב ג: כל העיר, ללא אילוצים
        for garden_row in gardens_records:
            garden_id = clean_value(garden_row.get("garden_id", ""))
            if garden_id and to_int(garden_row.get("capacity", DEFAULT_CAPACITY), default=DEFAULT_CAPACITY) > 0:
                candidate_rows.append(garden_row)
        fallback_stage = "city_wide_no_constraints"
    # ─────────────────────────────────────────────────────────────────────────

    filtering_summary = {
        "parent_id": parent_id,
        "parent_name": parent_row.get("parent_name", ""),
        "max_distance_km": parent_row.get("max_distance_km", ""),
        "preferred_distance_km": parent_row.get("preferred_distance_km", DEFAULT_PREFERRED_DISTANCE_KM),
        "total_gardens_in_database": total_gardens,
        "active_gardens_capacity_gt_0": len(active_gardens),
        "age_compatible_active_gardens": len(age_compatible_active_gardens),
        "within_normal_distance_candidates": len(within_normal_distance_gardens),
        "hard_constraint_passed_candidates": len(hard_constraint_passed_gardens),
        "manual_preferences_selected": " | ".join(manual_ids),
        "manual_preferences_included_regular_stage": " | ".join(manual_included),
        "manual_preferences_excluded_due_distance_regular_stage": " | ".join(manual_excluded_due_distance),
        "manual_preferences_excluded_due_age_regular_stage": " | ".join(manual_excluded_due_age),
        "manual_preferences_excluded_due_other_hard_constraint": " | ".join(manual_excluded_due_other_hard_constraint),
        "manual_preferences_not_found": " | ".join(manual_not_found),
        "hard_rejection_counts": " | ".join([f"{key}:{value}" for key, value in sorted(hard_rejection_counts.items())]),
        "filtering_stage": "hard_filter_capacity_age_distance_tzaharon_gender_religious_if_marked_hard",
        "fallback_stage": fallback_stage,
        "max_candidate_limit": MAX_CANDIDATE_GARDENS_PER_PARENT,
        "candidate_gardens_before_score_limit": len(candidate_rows),
        "candidate_gardens_final": 0
    }

    # סמן אילו גנים הגיעו דרך הגיבוי (ללא מרחק)
    fallback_ids = set()
    if fallback_stage:
        normal_ids = set(hard_constraint_passed_gardens)
        for r in candidate_rows:
            gid = clean_value(r.get("garden_id", ""))
            if gid and gid not in normal_ids:
                fallback_ids.add(gid)

    return candidate_rows, filtering_summary, fallback_ids


def apply_social_pairing_bonus(candidate_map, parents_records):
    parents_df = pd.DataFrame(parents_records)
    mutual_pairs = build_mutual_friend_pairs(parents_df)
    parent_rows = {
        clean_value(row.get("parent_id", "")): row
        for row in parents_records
    }

    social_bonus_rows = []

    for parent_a, parent_b in mutual_pairs:
        candidates_a = candidate_map.get(parent_a, {})
        candidates_b = candidate_map.get(parent_b, {})
        shared_gardens = sorted(set(candidates_a.keys()) & set(candidates_b.keys()))

        row_a = parent_rows.get(parent_a, {})
        row_b = parent_rows.get(parent_b, {})
        importance_a = get_importance(row_a, "importance_same_friend", default=1)
        importance_b = get_importance(row_b, "importance_same_friend", default=1)
        bonus_a = min(SOCIAL_FRIEND_MAX_BONUS, importance_a * SOCIAL_FRIEND_BONUS_PER_IMPORTANCE)
        bonus_b = min(SOCIAL_FRIEND_MAX_BONUS, importance_b * SOCIAL_FRIEND_BONUS_PER_IMPORTANCE)

        for garden_id in shared_gardens:
            candidates_a[garden_id]["parent_fit_score"] += bonus_a
            candidates_a[garden_id]["parent_fit_reasons"] += f" | mutual_friend_shared_candidate_bonus_with_{parent_b}"
            candidates_b[garden_id]["parent_fit_score"] += bonus_b
            candidates_b[garden_id]["parent_fit_reasons"] += f" | mutual_friend_shared_candidate_bonus_with_{parent_a}"

        social_bonus_rows.append({
            "parent_id_1": parent_a,
            "parent_id_2": parent_b,
            "shared_candidate_gardens": len(shared_gardens),
            "bonus_parent_1": bonus_a,
            "bonus_parent_2": bonus_b,
            "note": "soft bonus only; does not override hard constraints or capacity"
        })

    return social_bonus_rows


def build_preferences(parents_df, gardens_df):
    parent_preferences = {}
    garden_rankings = {}
    score_lookup = {}
    garden_priority_lookup = {}

    parent_preference_rows = []
    garden_ranking_rows = []
    candidate_filtering_rows = []

    parents_records = parents_df.to_dict("records")
    gardens_records = gardens_df.to_dict("records")
    all_garden_ids = set(clean_value(row.get("garden_id", "")) for row in gardens_records)

    parent_lookup = {
        clean_value(row.get("parent_id", "")): row
        for row in parents_records
    }

    candidate_map = {}
    filtering_by_parent = {}

    for parent_row in parents_records:
        parent_id = clean_value(parent_row.get("parent_id", ""))

        if not parent_id:
            continue

        candidate_garden_rows, filtering_summary, fallback_ids = get_candidate_gardens_for_parent(parent_row, gardens_records, all_garden_ids)
        candidates_for_parent = {}

        for garden_row in candidate_garden_rows:
            garden_id = clean_value(garden_row.get("garden_id", ""))

            if not garden_id:
                continue

            is_fallback = garden_id in fallback_ids
            fit_score, reasons, acceptable = calculate_parent_fit_score(parent_row, garden_row, skip_distance=is_fallback)

            if not acceptable:
                continue

            manual_rank = get_manual_preference_rank(parent_row, garden_id)
            distance_info = get_distance_info(parent_row, garden_row)

            if fit_score >= MIN_PARENT_SCORE_FOR_PREFERENCE or manual_rank is not None:
                candidates_for_parent[garden_id] = {
                    "garden_id": garden_id,
                    "garden_name": garden_row.get("garden_name", ""),
                    "manual_preference_rank": manual_rank if manual_rank is not None else 999,
                    "parent_fit_score": fit_score,
                    "parent_fit_reasons": " | ".join(reasons),
                    "distance_km": distance_info["distance_km"]
                }

        candidate_map[parent_id] = candidates_for_parent
        filtering_by_parent[parent_id] = filtering_summary

    social_bonus_rows = apply_social_pairing_bonus(candidate_map, parents_records)

    for parent_id, candidates_dict in candidate_map.items():
        parent_row = parent_lookup.get(parent_id, {})
        candidates = list(candidates_dict.values())

        candidates = sorted(
            candidates,
            key=lambda item: (
                -item["parent_fit_score"],
                item["manual_preference_rank"],
                999999 if item["distance_km"] is None else item["distance_km"],
                clean_value(item["garden_name"]),
                clean_value(item["garden_id"])
            )
        )

        candidates = candidates[:MAX_CANDIDATE_GARDENS_PER_PARENT]

        filtering_summary = filtering_by_parent.get(parent_id, {})
        filtering_summary["candidate_gardens_final"] = len(candidates)
        candidate_filtering_rows.append(filtering_summary)

        parent_preferences[parent_id] = [item["garden_id"] for item in candidates]

        for rank, item in enumerate(candidates, start=1):
            score_lookup[(parent_id, item["garden_id"])] = {
                "parent_fit_score": round(item["parent_fit_score"], 2),
                "parent_fit_reasons": item["parent_fit_reasons"],
                "manual_preference_rank": "" if item["manual_preference_rank"] == 999 else item["manual_preference_rank"]
            }

            parent_preference_rows.append({
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "rank": rank,
                "garden_id": item["garden_id"],
                "garden_name": item["garden_name"],
                "manual_preference_rank": "" if item["manual_preference_rank"] == 999 else item["manual_preference_rank"],
                "parent_fit_score": round(item["parent_fit_score"], 2),
                "distance_km": item["distance_km"],
                "parent_fit_reasons": item["parent_fit_reasons"]
            })

    for garden_row in gardens_records:
        garden_id = clean_value(garden_row.get("garden_id", ""))

        if not garden_id:
            continue

        candidates = []

        for parent_row in parents_records:
            parent_id = clean_value(parent_row.get("parent_id", ""))

            if not parent_id:
                continue

            fit_data = score_lookup.get((parent_id, garden_id))

            if not fit_data:
                continue

            priority_score, priority_reasons = calculate_garden_priority_score(
                parent_row,
                garden_row,
                fit_data["parent_fit_score"]
            )

            if priority_score < 0:
                continue

            garden_priority_lookup[(garden_id, parent_id)] = {
                "garden_priority_score": round(priority_score, 2),
                "garden_priority_reasons": " | ".join(priority_reasons)
            }

            candidates.append({
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "garden_priority_score": priority_score,
                "garden_priority_reasons": " | ".join(priority_reasons)
            })

        candidates = sorted(
            candidates,
            key=lambda item: (
                -item["garden_priority_score"],
                clean_value(item["parent_id"])
            )
        )

        garden_rankings[garden_id] = {
            item["parent_id"]: rank
            for rank, item in enumerate(candidates, start=1)
        }

        for rank, item in enumerate(candidates, start=1):
            garden_ranking_rows.append({
                "garden_id": garden_id,
                "garden_name": garden_row.get("garden_name", ""),
                "rank": rank,
                "parent_id": item["parent_id"],
                "parent_name": item["parent_name"],
                "child_name": item["child_name"],
                "garden_priority_score": round(item["garden_priority_score"], 2),
                "garden_priority_reasons": item["garden_priority_reasons"]
            })

    return (
        parent_preferences,
        garden_rankings,
        score_lookup,
        garden_priority_lookup,
        pd.DataFrame(parent_preference_rows),
        pd.DataFrame(garden_ranking_rows),
        pd.DataFrame(candidate_filtering_rows),
        pd.DataFrame(social_bonus_rows)
    )

def gale_shapley_with_capacities(parent_preferences, garden_rankings, capacities):
    free_parents = deque(parent_preferences.keys())
    next_proposal_index = {parent_id: 0 for parent_id in parent_preferences}
    proposal_counts = {parent_id: 0 for parent_id in parent_preferences}

    parent_match = {parent_id: None for parent_id in parent_preferences}
    garden_matches = {garden_id: [] for garden_id in capacities}

    proposal_log = []
    iteration = 0

    while free_parents:
        parent_id = free_parents.popleft()
        preferences = parent_preferences.get(parent_id, [])

        if next_proposal_index[parent_id] >= len(preferences):
            proposal_log.append({
                "iteration": iteration,
                "event_type": "no_more_preferences",
                "parent_id": parent_id,
                "proposal_number_for_parent": proposal_counts[parent_id],
                "garden_id": "",
                "decision": "unmatched_after_gale_shapley",
                "reason": "parent_has_no_more_candidate_gardens_within_normal_distance",
                "capacity": "",
                "matches_before": "",
                "matches_after": "",
                "rejected_in_this_step": ""
            })
            continue

        garden_id = preferences[next_proposal_index[parent_id]]
        next_proposal_index[parent_id] += 1
        proposal_counts[parent_id] += 1
        iteration += 1

        capacity = capacities.get(garden_id, 0)
        matches_before = list(garden_matches.get(garden_id, []))

        if capacity <= 0:
            proposal_log.append({
                "iteration": iteration,
                "event_type": "proposal",
                "parent_id": parent_id,
                "proposal_number_for_parent": proposal_counts[parent_id],
                "garden_id": garden_id,
                "decision": "rejected",
                "reason": "capacity_zero",
                "capacity": capacity,
                "matches_before": " | ".join(matches_before),
                "matches_after": " | ".join(matches_before),
                "rejected_in_this_step": parent_id
            })

            if next_proposal_index[parent_id] < len(preferences):
                free_parents.append(parent_id)

            continue

        current_candidates = matches_before + [parent_id]
        ranking = garden_rankings.get(garden_id, {})

        current_candidates = sorted(
            current_candidates,
            key=lambda candidate_parent: ranking.get(candidate_parent, 10**9)
        )

        accepted = current_candidates[:capacity]
        rejected = current_candidates[capacity:]

        garden_matches[garden_id] = accepted

        for accepted_parent in accepted:
            parent_match[accepted_parent] = garden_id

        if parent_id in accepted:
            decision = "accepted_temporarily"
            reason = "free_capacity" if len(matches_before) < capacity else "ranked_high_enough"
        else:
            decision = "rejected"
            reason = "garden_capacity_full_and_parent_ranked_lower"

        proposal_log.append({
            "iteration": iteration,
            "event_type": "proposal",
            "parent_id": parent_id,
            "proposal_number_for_parent": proposal_counts[parent_id],
            "garden_id": garden_id,
            "decision": decision,
            "reason": reason,
            "capacity": capacity,
            "matches_before": " | ".join(matches_before),
            "matches_after": " | ".join(accepted),
            "rejected_in_this_step": " | ".join(rejected)
        })

        for rejected_parent in rejected:
            parent_match[rejected_parent] = None

            if rejected_parent != parent_id:
                proposal_log.append({
                    "iteration": iteration,
                    "event_type": "replacement",
                    "parent_id": rejected_parent,
                    "proposal_number_for_parent": proposal_counts.get(rejected_parent, 0),
                    "garden_id": garden_id,
                    "decision": "rejected_after_replacement",
                    "reason": "garden_preferred_other_candidate",
                    "capacity": capacity,
                    "matches_before": " | ".join(matches_before),
                    "matches_after": " | ".join(accepted),
                    "rejected_in_this_step": " | ".join(rejected)
                })

            if next_proposal_index[rejected_parent] < len(parent_preferences.get(rejected_parent, [])):
                free_parents.append(rejected_parent)

    proposal_log_df = pd.DataFrame(proposal_log, columns=PROPOSAL_LOG_COLUMNS)

    return parent_match, garden_matches, proposal_log_df


def get_remaining_capacity(garden_matches, capacities, garden_id):
    return capacities.get(garden_id, 0) - len(garden_matches.get(garden_id, []))


def mandatory_nearby_assignment(parents_df, gardens_df, parent_match, garden_matches, capacities, score_lookup, garden_priority_lookup):
    parent_lookup = {
        clean_value(row.get("parent_id", "")): row
        for _, row in parents_df.iterrows()
    }

    mandatory_rows = []
    shortage_rows = []

    for parent_id, current_garden_id in list(parent_match.items()):
        if current_garden_id is not None:
            continue

        parent_row = parent_lookup.get(parent_id)

        if parent_row is None:
            continue

        candidates = []

        for _, garden_row in gardens_df.iterrows():
            garden_id = clean_value(garden_row.get("garden_id", ""))

            if not garden_id:
                continue

            if get_remaining_capacity(garden_matches, capacities, garden_id) <= 0:
                continue

            compatible, age_reason = age_is_compatible(parent_row, garden_row)

            if not compatible:
                continue

            distance_info = get_distance_info(parent_row, garden_row)

            if not distance_info["within_normal_distance"]:
                continue

            fit_score, fit_reasons, acceptable = calculate_parent_fit_score(parent_row, garden_row)

            if not acceptable:
                continue

            candidates.append({
                "garden_id": garden_id,
                "garden_name": garden_row.get("garden_name", ""),
                "fit_score": fit_score,
                "fit_reasons": " | ".join(fit_reasons),
                "distance_km": distance_info["distance_km"],
                "max_distance_km": distance_info["max_distance_km"],
                "same_neighborhood": distance_info["same_neighborhood"]
            })

        # גיבוי: אם לא נמצאו מועמדים (למשל הורה ללא קואורדינטות ושכונה לא מזוהה)
        # חפש ברחבי העיר לפי גיל בלבד
        if not candidates:
            for _, garden_row in gardens_df.iterrows():
                garden_id = clean_value(garden_row.get("garden_id", ""))
                if not garden_id:
                    continue
                if get_remaining_capacity(garden_matches, capacities, garden_id) <= 0:
                    continue
                compatible, _ = age_is_compatible(parent_row, garden_row)
                if not compatible:
                    continue
                fit_score, fit_reasons, _ = calculate_parent_fit_score(parent_row, garden_row, skip_distance=True)
                candidates.append({
                    "garden_id": garden_id,
                    "garden_name": garden_row.get("garden_name", ""),
                    "fit_score": fit_score,
                    "fit_reasons": " | ".join(fit_reasons) + " | city_wide_fallback",
                    "distance_km": None,
                    "max_distance_km": None,
                    "same_neighborhood": False
                })

        # גיבוי אחרון: ללא אילוצי גיל
        if not candidates:
            for _, garden_row in gardens_df.iterrows():
                garden_id = clean_value(garden_row.get("garden_id", ""))
                if not garden_id:
                    continue
                if get_remaining_capacity(garden_matches, capacities, garden_id) <= 0:
                    continue
                candidates.append({
                    "garden_id": garden_id,
                    "garden_name": garden_row.get("garden_name", ""),
                    "fit_score": 0,
                    "fit_reasons": "last_resort_fallback",
                    "distance_km": None,
                    "max_distance_km": None,
                    "same_neighborhood": False
                })

        candidates = sorted(
            candidates,
            key=lambda item: (
                -item["fit_score"],
                999999 if item["distance_km"] is None else item["distance_km"],
                clean_value(item["garden_name"])
            )
        )

        if candidates:
            selected = candidates[0]
            selected_garden_id = selected["garden_id"]

            garden_matches.setdefault(selected_garden_id, []).append(parent_id)
            parent_match[parent_id] = selected_garden_id

            score_lookup[(parent_id, selected_garden_id)] = {
                "parent_fit_score": round(selected["fit_score"], 2),
                "parent_fit_reasons": selected["fit_reasons"] + " | mandatory_nearby_assignment",
                "manual_preference_rank": get_manual_preference_rank(parent_row, selected_garden_id) or ""
            }

            garden_priority_lookup[(selected_garden_id, parent_id)] = {
                "garden_priority_score": "",
                "garden_priority_reasons": "mandatory assignment into remaining nearby capacity"
            }

            mandatory_rows.append({
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "assigned_garden_id": selected_garden_id,
                "assigned_garden_name": selected["garden_name"],
                "assignment_status": "matched_mandatory_nearby",
                "fit_score": round(selected["fit_score"], 2),
                "distance_km": selected["distance_km"],
                "max_distance_km": selected["max_distance_km"],
                "reason": "unmatched_after_gale_shapley_assigned_to_best_available_nearby_garden"
            })
        else:
            shortage_rows.append({
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "home_neighborhood": parent_row.get("home_neighborhood", ""),
                "home_address": parent_row.get("home_address", ""),
                "max_distance_km": parent_row.get("max_distance_km", ""),
                "status": "requires_planning_review",
                "reason": "no_available_capacity_within_normal_distance"
            })

    return pd.DataFrame(mandatory_rows), pd.DataFrame(shortage_rows)


def apply_manual_exception_after_all_close_solution(parents_df, gardens_df, parent_match, garden_matches, capacities, score_lookup, garden_priority_lookup, shortage_rows_df):
    if not shortage_rows_df.empty:
        return pd.DataFrame([{
            "action": "manual_exception_stage_skipped",
            "reason": "there_are_children_without_nearby_solution_so_manual_far_preferences_cannot_use_remaining_capacity",
            "planning_review_count": len(shortage_rows_df)
        }])

    parent_lookup = {
        clean_value(row.get("parent_id", "")): row
        for _, row in parents_df.iterrows()
    }

    garden_lookup = {
        clean_value(row.get("garden_id", "")): row
        for _, row in gardens_df.iterrows()
    }

    log_rows = []

    for parent_id, current_garden_id in list(parent_match.items()):
        if current_garden_id is None:
            continue

        parent_row = parent_lookup.get(parent_id)

        if parent_row is None:
            continue

        manual_ids = [garden_id for garden_id in get_manual_preference_ids(parent_row) if garden_id]

        for manual_garden_id in manual_ids:
            if manual_garden_id == current_garden_id:
                log_rows.append({
                    "parent_id": parent_id,
                    "current_garden_id": current_garden_id,
                    "manual_garden_id": manual_garden_id,
                    "decision": "no_change",
                    "reason": "already_assigned_to_manual_preference",
                    "distance_km": "",
                    "max_distance_km": parent_row.get("max_distance_km", ""),
                    "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", "")
                })
                break

            garden_row = garden_lookup.get(manual_garden_id)

            if garden_row is None:
                log_rows.append({
                    "parent_id": parent_id,
                    "current_garden_id": current_garden_id,
                    "manual_garden_id": manual_garden_id,
                    "decision": "not_applied",
                    "reason": "manual_garden_not_found",
                    "distance_km": "",
                    "max_distance_km": parent_row.get("max_distance_km", ""),
                    "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", "")
                })
                continue

            if get_remaining_capacity(garden_matches, capacities, manual_garden_id) <= 0:
                log_rows.append({
                    "parent_id": parent_id,
                    "current_garden_id": current_garden_id,
                    "manual_garden_id": manual_garden_id,
                    "decision": "not_applied",
                    "reason": "manual_garden_has_no_remaining_capacity",
                    "distance_km": get_distance_info(parent_row, garden_row)["distance_km"],
                    "max_distance_km": parent_row.get("max_distance_km", ""),
                    "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", "")
                })
                continue

            distance_info = get_distance_info(parent_row, garden_row)

            if distance_info["within_normal_distance"]:
                allowed = True
                reason = "manual_preference_within_normal_distance_and_capacity_left"
            else:
                allowed, reason = can_use_manual_exception(parent_row, garden_row)

            if not allowed:
                log_rows.append({
                    "parent_id": parent_id,
                    "current_garden_id": current_garden_id,
                    "manual_garden_id": manual_garden_id,
                    "decision": "not_applied",
                    "reason": reason,
                    "distance_km": distance_info["distance_km"],
                    "max_distance_km": parent_row.get("max_distance_km", ""),
                    "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", "")
                })
                continue

            if current_garden_id in garden_matches and parent_id in garden_matches[current_garden_id]:
                garden_matches[current_garden_id].remove(parent_id)

            garden_matches.setdefault(manual_garden_id, []).append(parent_id)
            parent_match[parent_id] = manual_garden_id

            fit_score, fit_reasons, _ = calculate_parent_fit_score_for_manual_exception(parent_row, garden_row)

            score_lookup[(parent_id, manual_garden_id)] = {
                "parent_fit_score": round(fit_score, 2),
                "parent_fit_reasons": " | ".join(fit_reasons) + " | final_manual_exception_stage",
                "manual_preference_rank": get_manual_preference_rank(parent_row, manual_garden_id) or ""
            }

            garden_priority_lookup[(manual_garden_id, parent_id)] = {
                "garden_priority_score": "",
                "garden_priority_reasons": "manual exception applied after all children had nearby solution"
            }

            log_rows.append({
                "parent_id": parent_id,
                "current_garden_id": current_garden_id,
                "manual_garden_id": manual_garden_id,
                "decision": "applied",
                "reason": reason,
                "distance_km": distance_info["distance_km"],
                "max_distance_km": parent_row.get("max_distance_km", ""),
                "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", "")
            })
            break

    return pd.DataFrame(log_rows)


def calculate_parent_fit_score_for_manual_exception(parent_row, garden_row):
    score = 0.0
    reasons = []

    compatible, age_reason = age_is_compatible(parent_row, garden_row)

    if not compatible:
        return 0.0, [age_reason], False

    reasons.append(age_reason)

    distance_info = get_distance_info(parent_row, garden_row)
    distance = distance_info["distance_km"]
    max_distance = distance_info["max_distance_km"]

    if distance_info["within_normal_distance"]:
        reasons.append("within_normal_distance")
        if distance is not None and max_distance is not None and max_distance > 0:
            score += max((1 - distance / max_distance) * get_importance(parent_row, "importance_distance") * 20, 0)
    else:
        reasons.append("manual_exception_distance")

    manual_rank = get_manual_preference_rank(parent_row, garden_row.get("garden_id"))

    if manual_rank is not None:
        score += {1: 300, 2: 250, 3: 200}.get(manual_rank, 0)
        reasons.append(f"manual_preference_{manual_rank}")

    if contains_match(parent_row.get("preferred_activity_language"), garden_row.get("activity_language")):
        score += get_importance(parent_row, "importance_activity_language") * 15
        reasons.append("activity_language_match")

    if contains_match(parent_row.get("preferred_education_type"), garden_row.get("education_type")):
        score += get_importance(parent_row, "importance_education_type") * 15
        reasons.append("education_type_match")

    if contains_match(parent_row.get("preferred_religious_orientation"), garden_row.get("religious_orientation")):
        score += get_importance(parent_row, "importance_religious_orientation") * 15
        reasons.append("religious_orientation_match")

    if contains_match(parent_row.get("preferred_sector"), garden_row.get("sector")):
        score += get_importance(parent_row, "importance_sector") * 10
        reasons.append("sector_match")

    return score, reasons, True


def build_stability_check(parent_preferences, garden_rankings, parent_match, garden_matches, capacities, parents_df, gardens_df):
    parent_rankings = {
        parent_id: {
            garden_id: rank
            for rank, garden_id in enumerate(preferences, start=1)
        }
        for parent_id, preferences in parent_preferences.items()
    }

    garden_name_lookup = {
        clean_value(row.get("garden_id", "")): row.get("garden_name", "")
        for _, row in gardens_df.iterrows()
    }

    parent_name_lookup = {
        clean_value(row.get("parent_id", "")): row.get("parent_name", "")
        for _, row in parents_df.iterrows()
    }

    blocking_pairs = []

    for parent_id, preferences in parent_preferences.items():
        current_garden_id = parent_match.get(parent_id)
        parent_ranking = parent_rankings.get(parent_id, {})

        if current_garden_id is None:
            preferred_gardens_to_check = preferences
            current_parent_rank = "unmatched"
        else:
            current_rank = parent_ranking.get(current_garden_id)
            current_parent_rank = current_rank if current_rank is not None else "outside_regular_preference_list"

            if current_rank is None:
                preferred_gardens_to_check = preferences
            else:
                preferred_gardens_to_check = preferences[:current_rank - 1]

        for preferred_garden_id in preferred_gardens_to_check:
            capacity = capacities.get(preferred_garden_id, 0)

            if capacity <= 0:
                continue

            matched_parents = garden_matches.get(preferred_garden_id, [])
            garden_ranking = garden_rankings.get(preferred_garden_id, {})
            parent_rank_at_garden = garden_ranking.get(parent_id, 10**9)

            if parent_rank_at_garden == 10**9:
                continue

            if len(matched_parents) < capacity:
                blocking_pairs.append({
                    "parent_id": parent_id,
                    "parent_name": parent_name_lookup.get(parent_id, ""),
                    "current_garden_id": current_garden_id or "",
                    "current_garden_name": garden_name_lookup.get(current_garden_id, "") if current_garden_id else "",
                    "preferred_garden_id": preferred_garden_id,
                    "preferred_garden_name": garden_name_lookup.get(preferred_garden_id, ""),
                    "reason": "parent_prefers_garden_and_garden_has_free_capacity",
                    "current_parent_preference_rank": current_parent_rank,
                    "preferred_garden_rank_for_parent": parent_ranking.get(preferred_garden_id, ""),
                    "parent_rank_at_preferred_garden": parent_rank_at_garden,
                    "worst_current_parent_at_garden": "",
                    "worst_current_parent_rank_at_garden": ""
                })
                continue

            worst_current_parent = max(
                matched_parents,
                key=lambda candidate_parent: garden_ranking.get(candidate_parent, 10**9)
            )

            worst_current_parent_rank = garden_ranking.get(worst_current_parent, 10**9)

            if parent_rank_at_garden < worst_current_parent_rank:
                blocking_pairs.append({
                    "parent_id": parent_id,
                    "parent_name": parent_name_lookup.get(parent_id, ""),
                    "current_garden_id": current_garden_id or "",
                    "current_garden_name": garden_name_lookup.get(current_garden_id, "") if current_garden_id else "",
                    "preferred_garden_id": preferred_garden_id,
                    "preferred_garden_name": garden_name_lookup.get(preferred_garden_id, ""),
                    "reason": "parent_and_garden_prefer_each_other_over_current_matching",
                    "current_parent_preference_rank": current_parent_rank,
                    "preferred_garden_rank_for_parent": parent_ranking.get(preferred_garden_id, ""),
                    "parent_rank_at_preferred_garden": parent_rank_at_garden,
                    "worst_current_parent_at_garden": worst_current_parent,
                    "worst_current_parent_rank_at_garden": worst_current_parent_rank
                })

    blocking_pairs_df = pd.DataFrame(blocking_pairs)

    stability_summary_df = pd.DataFrame([
        {
            "metric": "stability_scope",
            "value": "regular_gale_shapley_candidate_lists_after_hard_distance_filter"
        },
        {
            "metric": "stable_matching",
            "value": "yes" if blocking_pairs_df.empty else "no"
        },
        {
            "metric": "blocking_pairs_found",
            "value": len(blocking_pairs_df)
        },
        {
            "metric": "meaning",
            "value": "No blocking pairs means no parent and garden in the regular candidate lists prefer each other over the final matching."
        }
    ])

    return stability_summary_df, blocking_pairs_df


def build_distance_audit(parents_df, gardens_df, parent_match, assignment_source):
    parent_lookup = {
        clean_value(row.get("parent_id", "")): row
        for _, row in parents_df.iterrows()
    }

    garden_lookup = {
        clean_value(row.get("garden_id", "")): row
        for _, row in gardens_df.iterrows()
    }

    rows = []

    for parent_id, garden_id in parent_match.items():
        parent_row = parent_lookup.get(parent_id)

        if parent_row is None:
            continue

        if garden_id is None:
            rows.append({
                "parent_id": parent_id,
                "assignment_source": assignment_source.get(parent_id, "requires_planning_review"),
                "garden_id": "",
                "distance_km": "",
                "max_distance_km": parent_row.get("max_distance_km", ""),
                "within_normal_distance": "no",
                "allow_manual_far_preference": parent_row.get("allow_manual_far_preference", ""),
                "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", ""),
                "distance_status": "no_assignment_due_planning_review"
            })
            continue

        garden_row = garden_lookup.get(garden_id)

        if garden_row is None:
            continue

        distance_info = get_distance_info(parent_row, garden_row)

        rows.append({
            "parent_id": parent_id,
            "assignment_source": assignment_source.get(parent_id, ""),
            "garden_id": garden_id,
            "distance_km": distance_info["distance_km"],
            "max_distance_km": distance_info["max_distance_km"],
            "within_normal_distance": "yes" if distance_info["within_normal_distance"] else "no",
            "allow_manual_far_preference": parent_row.get("allow_manual_far_preference", ""),
            "max_manual_exception_distance_km": parent_row.get("max_manual_exception_distance_km", ""),
            "distance_status": distance_info["distance_reason"]
        })

    return pd.DataFrame(rows)


def build_preference_satisfaction(parents_df, gardens_df, parent_match, assignment_source):
    garden_lookup = {
        clean_value(row.get("garden_id", "")): row
        for _, row in gardens_df.iterrows()
    }
    rows = []
    for _, parent_row in parents_df.iterrows():
        parent_id = clean_value(parent_row.get("parent_id", ""))
        garden_id = parent_match.get(parent_id)
        garden_row = garden_lookup.get(garden_id, {}) if garden_id is not None else {}
        manual_rank = get_manual_preference_rank(parent_row, garden_id) if garden_id else None
        distance_info = get_distance_info(parent_row, garden_row) if garden_id and len(garden_row) else {
            "distance_km": "", "max_distance_km": parent_row.get("max_distance_km", ""), "within_normal_distance": False
        }
        if garden_id is None:
            category = "requires_planning_review"
        elif manual_rank == 1:
            category = "got_first_manual_preference"
        elif manual_rank == 2:
            category = "got_second_manual_preference"
        elif manual_rank == 3:
            category = "got_third_manual_preference"
        else:
            category = "matched_outside_manual_preferences"

        rows.append({
            "parent_id": parent_id,
            "child_name": parent_row.get("child_name", ""),
            "matched_garden_id": garden_id or "",
            "manual_preference_rank_received": manual_rank or "",
            "satisfaction_category": category,
            "assignment_source": assignment_source.get(parent_id, ""),
            "distance_km": distance_info.get("distance_km", ""),
            "max_distance_km": distance_info.get("max_distance_km", ""),
            "within_normal_distance": "yes" if distance_info.get("within_normal_distance") else "no"
        })
    return pd.DataFrame(rows)


def build_garden_demand_ranking(parents_df, final_df):
    demand = {}
    for _, row in parents_df.iterrows():
        for rank, column in [(1, "preferred_garden_1"), (2, "preferred_garden_2"), (3, "preferred_garden_3")]:
            garden_id = clean_value(row.get(column, ""))
            if not garden_id:
                continue
            demand.setdefault(garden_id, {"garden_id": garden_id, "rank_1_requests": 0, "rank_2_requests": 0, "rank_3_requests": 0, "total_manual_requests": 0})
            demand[garden_id][f"rank_{rank}_requests"] += 1
            demand[garden_id]["total_manual_requests"] += 1

    assigned_counts = {}
    if not final_df.empty and "matched_garden_id" in final_df.columns:
        for garden_id, count in final_df["matched_garden_id"].value_counts(dropna=True).items():
            garden_id = clean_value(garden_id)
            if garden_id:
                assigned_counts[garden_id] = int(count)

    rows = []
    for garden_id, data in demand.items():
        data = dict(data)
        data["assigned_count"] = assigned_counts.get(garden_id, 0)
        rows.append(data)

    return pd.DataFrame(rows).sort_values(
        by=["rank_1_requests", "total_manual_requests", "assigned_count"],
        ascending=[False, False, False]
    ) if rows else pd.DataFrame(columns=["garden_id", "rank_1_requests", "rank_2_requests", "rank_3_requests", "total_manual_requests", "assigned_count"])


def build_demand_by_criteria(parents_df):
    rows = []
    total = len(parents_df)
    criteria_columns = [
        "max_distance_km", "needs_tzaharon", "preferred_religious_orientation",
        "preferred_religious_substream", "preferred_gender_composition", "needs_friday",
        "needs_protected_space", "preferred_activity_language", "preferred_education_type",
        "child_birth_month"
    ]
    for column in criteria_columns:
        if column not in parents_df.columns:
            continue
        counts = parents_df[column].apply(clean_value).replace("", "not_specified").value_counts()
        for value, count in counts.items():
            rows.append({
                "criterion": column,
                "value": value,
                "parents_count": int(count),
                "parents_percent": round((int(count) / total) * 100, 2) if total else 0
            })
    return pd.DataFrame(rows)


def build_birth_month_audit(parents_df, final_df, distance_audit_df):
    if "child_birth_month" not in parents_df.columns:
        return pd.DataFrame([{"note": "child_birth_month column missing"}])

    data = parents_df[["parent_id", "child_birth_month"]].copy()
    data["parent_id"] = data["parent_id"].apply(clean_value)
    data["child_birth_month"] = data["child_birth_month"].apply(clean_value)
    data = data[data["child_birth_month"] != ""]

    if data.empty:
        return pd.DataFrame([{"note": "child_birth_month values missing"}])

    sat_cols = ["parent_id", "manual_preference_rank_received", "satisfaction_category"]
    if not final_df.empty:
        final_small = final_df[["parent_id", "match_status"]].copy()
        final_small["parent_id"] = final_small["parent_id"].apply(clean_value)
        data = data.merge(final_small, on="parent_id", how="left")

    if not distance_audit_df.empty:
        dist = distance_audit_df[["parent_id", "distance_km"]].copy()
        dist["parent_id"] = dist["parent_id"].apply(clean_value)
        data = data.merge(dist, on="parent_id", how="left")

    rows = []
    for month, group in data.groupby("child_birth_month"):
        distances = pd.to_numeric(group.get("distance_km", pd.Series(dtype=float)), errors="coerce")
        rows.append({
            "child_birth_month": month,
            "parents_count": len(group),
            "matched_count": int((group.get("match_status", pd.Series(dtype=str)) == "matched").sum()) if "match_status" in group.columns else "",
            "avg_distance_km": round(float(distances.mean()), 3) if not distances.dropna().empty else "",
            "max_distance_km": round(float(distances.max()), 3) if not distances.dropna().empty else ""
        })
    return pd.DataFrame(rows).sort_values("child_birth_month")


def build_social_pairing_check(parents_df, parent_match):
    parent_lookup = {
        clean_value(row.get("parent_id", "")): row
        for _, row in parents_df.iterrows()
    }
    requests = {
        parent_id: set(friend_id for friend_id in get_parent_requested_friend_ids(row) if friend_id)
        for parent_id, row in parent_lookup.items()
    }
    rows = []
    seen_mutual_pairs = set()
    for parent_id, requested_ids in requests.items():
        for friend_id in requested_ids:
            mutual = friend_id in requests and parent_id in requests.get(friend_id, set())
            pair_key = tuple(sorted([parent_id, friend_id]))
            if mutual and pair_key in seen_mutual_pairs:
                continue
            if mutual:
                seen_mutual_pairs.add(pair_key)
            parent_garden = parent_match.get(parent_id)
            friend_garden = parent_match.get(friend_id)
            rows.append({
                "parent_id": parent_id,
                "friend_parent_id": friend_id,
                "request_type": "mutual" if mutual else "one_sided",
                "importance_same_friend": parent_lookup.get(parent_id, {}).get("importance_same_friend", ""),
                "willing_to_trade_distance_for_friend": parent_lookup.get(parent_id, {}).get("willing_to_trade_distance_for_friend", ""),
                "parent_matched_garden_id": parent_garden or "",
                "friend_matched_garden_id": friend_garden or "",
                "same_garden_result": "yes" if parent_garden and parent_garden == friend_garden else "no",
                "note": "soft preference only" if mutual else "ignored because friend did not request back"
            })
    return pd.DataFrame(rows) if rows else pd.DataFrame([{"note": "no friend requests found"}])


def create_outputs(
    parents_df,
    gardens_df,
    parent_match,
    garden_matches,
    score_lookup,
    garden_priority_lookup,
    parent_preferences_df,
    garden_rankings_df,
    candidate_filtering_df,
    proposal_log_df,
    stability_summary_df,
    blocking_pairs_df,
    capacities,
    mandatory_assignment_df,
    planning_review_df,
    manual_exception_log_df,
    distance_audit_df,
    assignment_source,
    social_bonus_df
):
    garden_name_lookup = {
        clean_value(row.get("garden_id", "")): row.get("garden_name", "")
        for _, row in gardens_df.iterrows()
    }

    parent_lookup = {
        clean_value(row.get("parent_id", "")): row.to_dict()
        for _, row in parents_df.iterrows()
    }

    final_rows = []
    shortage_rows = []

    for parent_id, matched_garden_id in parent_match.items():
        parent_row = parent_lookup.get(parent_id, {})
        source = assignment_source.get(parent_id, "")

        if matched_garden_id is None:
            row = {
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "match_status": "requires_planning_review",
                "assignment_source": "requires_planning_review",
                "matched_garden_id": "",
                "matched_garden_name": "",
                "parent_fit_score": "",
                "parent_fit_reasons": "",
                "garden_priority_score": "",
                "garden_priority_reasons": ""
            }

            final_rows.append(row)
            shortage_rows.append(row)

        else:
            fit_data = score_lookup.get((parent_id, matched_garden_id), {})
            priority_data = garden_priority_lookup.get((matched_garden_id, parent_id), {})

            final_rows.append({
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "match_status": "matched",
                "assignment_source": source,
                "matched_garden_id": matched_garden_id,
                "matched_garden_name": garden_name_lookup.get(matched_garden_id, ""),
                "parent_fit_score": fit_data.get("parent_fit_score", ""),
                "parent_fit_reasons": fit_data.get("parent_fit_reasons", ""),
                "garden_priority_score": priority_data.get("garden_priority_score", ""),
                "garden_priority_reasons": priority_data.get("garden_priority_reasons", "")
            })

    assignment_rows = []

    for garden_id, parent_ids in garden_matches.items():
        for parent_id in parent_ids:
            parent_row = parent_lookup.get(parent_id, {})
            fit_data = score_lookup.get((parent_id, garden_id), {})
            priority_data = garden_priority_lookup.get((garden_id, parent_id), {})

            assignment_rows.append({
                "garden_id": garden_id,
                "garden_name": garden_name_lookup.get(garden_id, ""),
                "parent_id": parent_id,
                "parent_name": parent_row.get("parent_name", ""),
                "child_name": parent_row.get("child_name", ""),
                "assignment_source": assignment_source.get(parent_id, ""),
                "parent_fit_score": fit_data.get("parent_fit_score", ""),
                "garden_priority_score": priority_data.get("garden_priority_score", "")
            })

    capacity_rows = []

    for garden_id, capacity in capacities.items():
        assigned_count = len(garden_matches.get(garden_id, []))

        capacity_rows.append({
            "garden_id": garden_id,
            "garden_name": garden_name_lookup.get(garden_id, ""),
            "capacity": capacity,
            "assigned_count": assigned_count,
            "remaining_capacity": max(capacity - assigned_count, 0),
            "is_full": "yes" if assigned_count >= capacity and capacity > 0 else "no"
        })

    final_df = pd.DataFrame(final_rows)
    assignment_df = pd.DataFrame(assignment_rows)
    capacity_df = pd.DataFrame(capacity_rows)

    parents_total = len(parent_match)
    matched_total = int((final_df["match_status"] == "matched").sum()) if not final_df.empty else 0
    shortage_total = int((final_df["match_status"] == "requires_planning_review").sum()) if not final_df.empty else 0

    gardens_total = len(gardens_df)
    active_gardens_total = sum(1 for capacity in capacities.values() if capacity > 0)

    candidate_pairs_total = len(parent_preferences_df)
    naive_pairs_total = parents_total * gardens_total
    active_pairs_total = parents_total * active_gardens_total

    average_candidates_per_parent = round(candidate_pairs_total / parents_total, 2) if parents_total else 0

    if not stability_summary_df.empty:
        stable_matching_value = str(
            stability_summary_df.loc[
                stability_summary_df["metric"] == "stable_matching",
                "value"
            ].iloc[0]
        )

        blocking_pairs_value = int(
            stability_summary_df.loc[
                stability_summary_df["metric"] == "blocking_pairs_found",
                "value"
            ].iloc[0]
        )
    else:
        stable_matching_value = ""
        blocking_pairs_value = ""

    summary_df = pd.DataFrame([
        {"metric": "parents_total", "value": parents_total},
        {"metric": "parents_matched_total", "value": matched_total},
        {"metric": "matched_by_gale_shapley", "value": sum(1 for v in assignment_source.values() if v == "gale_shapley")},
        {"metric": "matched_by_mandatory_nearby_assignment", "value": sum(1 for v in assignment_source.values() if v == "mandatory_nearby_assignment")},
        {"metric": "matched_by_final_manual_exception", "value": sum(1 for v in assignment_source.values() if v == "final_manual_exception")},
        {"metric": "requires_planning_review", "value": shortage_total},
        {"metric": "gardens_total", "value": gardens_total},
        {"metric": "active_gardens_capacity_gt_0", "value": active_gardens_total},
        {"metric": "gardens_used", "value": len(set(assignment_df["garden_id"])) if not assignment_df.empty else 0},
        {"metric": "candidate_pairs_after_hard_distance_filter", "value": candidate_pairs_total},
        {"metric": "average_candidates_per_parent", "value": average_candidates_per_parent},
        {"metric": "blocking_pairs_found", "value": blocking_pairs_value},
        {"metric": "stable_matching", "value": stable_matching_value}
    ])

    complexity_df = pd.DataFrame([
        {
            "item": "naive_pairwise_comparisons",
            "value": naive_pairs_total,
            "explanation": "n*m: every parent compared with every garden"
        },
        {
            "item": "active_capacity_pairwise_comparisons",
            "value": active_pairs_total,
            "explanation": "parents compared only with gardens that have capacity > 0"
        },
        {
            "item": "candidate_pairs_after_hard_distance_filter",
            "value": candidate_pairs_total,
            "explanation": "actual parent-garden pairs kept after capacity, age and normal-distance filtering"
        },
        {
            "item": "average_candidates_per_parent",
            "value": average_candidates_per_parent,
            "explanation": "average k value after filtering"
        },
        {
            "item": "theoretical_before",
            "value": "O(n*m*c)",
            "explanation": "without blocking/filtering"
        },
        {
            "item": "theoretical_after",
            "value": "O(n*k*c)",
            "explanation": "after blocking, where k is relevant nearby gardens per parent"
        },
        {
            "item": "max_candidate_gardens_per_parent",
            "value": MAX_CANDIDATE_GARDENS_PER_PARENT,
            "explanation": "upper limit used by the current implementation during the regular Gale-Shapley stage"
        }
    ])

    # ── גיליון 1: טבלת שיבוצים ──────────────────────────────────────
    garden_neighborhood_lookup = {
        clean_value(row.get("garden_id", "")): row.get("neighborhood", "")
        for _, row in gardens_df.iterrows()
    }
    garden_address_lookup = {
        clean_value(row.get("garden_id", "")): row.get("address", "")
        for _, row in gardens_df.iterrows()
    }

    def clean_garden_id_field(val):
        try:
            return str(int(float(val))) if val not in (None, "", float('nan')) else ""
        except (ValueError, TypeError):
            return clean_value(val)

    pref_rank_lookup = {}
    for _, row in parents_df.iterrows():
        pid = clean_value(row.get("parent_id", ""))
        gid = parent_match.get(pid)
        p1 = clean_garden_id_field(row.get("preferred_garden_1", ""))
        p2 = clean_garden_id_field(row.get("preferred_garden_2", ""))
        p3 = clean_garden_id_field(row.get("preferred_garden_3", ""))
        if gid == p1:
            rank = "עדיפות 1"
        elif gid == p2:
            rank = "עדיפות 2"
        elif gid == p3:
            rank = "עדיפות 3"
        else:
            rank = "מחוץ לרשימה" if gid else "לא שובץ"
        pref_rank_lookup[pid] = rank

    matching_rows = []
    for _, row in parents_df.iterrows():
        pid = clean_value(row.get("parent_id", ""))
        gid = parent_match.get(pid)
        matching_rows.append({
            "שם הורה": clean_value(row.get("parent_name", "")),
            "שם ילד": clean_value(row.get("child_name", "")),
            "גיל (חודשים)": clean_value(row.get("child_age_months", "")),
            "שכונת המגורים": clean_value(row.get("home_neighborhood", "")),
            "גן משובץ": garden_name_lookup.get(gid, "") if gid else "לא שובץ",
            "שכונת הגן": garden_neighborhood_lookup.get(gid, "") if gid else "",
            "כתובת הגן": garden_address_lookup.get(gid, "") if gid else "",
            "עדיפות שקיבל": pref_rank_lookup.get(pid, ""),
        })
    matching_clean_df = pd.DataFrame(matching_rows)

    # ── גיליון 2: סטטוס גנים ──────────────────────────────────────
    garden_demand = {}
    for _, row in parents_df.iterrows():
        for field in ["preferred_garden_1", "preferred_garden_2", "preferred_garden_3"]:
            gid = clean_garden_id_field(row.get(field, ""))
            if gid:
                garden_demand[gid] = garden_demand.get(gid, 0) + 1

    garden_rows = []
    for _, row in gardens_df.iterrows():
        gid = clean_value(row.get("garden_id", ""))
        if not gid:
            continue
        cap = to_int(row.get("capacity", DEFAULT_CAPACITY), default=DEFAULT_CAPACITY)
        assigned = len(garden_matches.get(gid, []))
        garden_rows.append({
            "שם גן": clean_value(row.get("garden_name", "")),
            "שכונה": clean_value(row.get("neighborhood", "")),
            "קיבולת": cap,
            "ילדים שובצו": assigned,
            "מקומות פנויים": max(cap - assigned, 0),
            "ביקוש (הורים שביקשו)": garden_demand.get(gid, 0),
        })
    gardens_clean_df = pd.DataFrame(garden_rows).sort_values("ילדים שובצו", ascending=False)

    # ── דוחות נוספים לקובץ המפורט ─────────────────────────────────
    preference_satisfaction_df = build_preference_satisfaction(parents_df, gardens_df, parent_match, assignment_source)
    garden_demand_ranking_df = build_garden_demand_ranking(parents_df, final_df)
    demand_by_criteria_df = build_demand_by_criteria(parents_df)
    birth_month_audit_df = build_birth_month_audit(parents_df, final_df, distance_audit_df)
    social_pairing_df = build_social_pairing_check(parents_df, parent_match)

    # קובץ קצר לאתר — נוח לקריאה מהירה
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        matching_clean_df.to_excel(writer, sheet_name="טבלת שיבוצים", index=False)
        gardens_clean_df.to_excel(writer, sheet_name="סטטוס גנים", index=False)

    # קובץ מפורט לדוח/בדיקה — כל גיליונות האלגוריתם
    full_output_sheets = [
        ("Summary", summary_df),
        ("Final_Matching", final_df),
        ("Garden_Assignments", assignment_df),
        ("Stability_Check", stability_summary_df),
        ("Blocking_Pairs", blocking_pairs_df),
        ("Complexity", complexity_df),
        ("Capacity_Utilization", capacity_df),
        ("Preference_Satisfaction", preference_satisfaction_df),
        ("Social_Pairing_Check", social_pairing_df),
        ("Demand_By_Criteria", demand_by_criteria_df),
        ("Garden_Demand_Ranking", garden_demand_ranking_df),
        ("Birth_Month_Audit", birth_month_audit_df),
        ("Social_Bonus", social_bonus_df),
        ("Mandatory_Assignment", mandatory_assignment_df),
        ("Planning_Review", planning_review_df),
        ("Manual_Exception_Log", manual_exception_log_df),
        ("Distance_Audit", distance_audit_df),
        ("Proposal_Log", proposal_log_df),
        ("Candidate_Filtering", candidate_filtering_df),
        ("Parent_Preferences", parent_preferences_df),
        ("Garden_Rankings", garden_rankings_df),
    ]

    with pd.ExcelWriter(FULL_OUTPUT_FILE, engine="openpyxl") as writer:
        for sheet_name, df in full_output_sheets:
            if df is None:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Done | Created {OUTPUT_FILE}")
    print(f"Done | Created {FULL_OUTPUT_FILE}")
    print(f"Parents: {parents_total}")
    print(f"Matched total: {matched_total}")
    print(f"Requires planning review: {shortage_total}")
    print(f"Proposal log rows: {len(proposal_log_df)}")
    print(f"Blocking pairs found: {len(blocking_pairs_df)}")
    print(f"Average candidates per parent: {average_candidates_per_parent}")


def main():
    gardens_df = ensure_gardens_input_file()

    if not os.path.exists(PARENTS_INPUT_FILE):
        raise FileNotFoundError(f"Missing {PARENTS_INPUT_FILE}")

    parents_df = pd.read_excel(PARENTS_INPUT_FILE)
    validate_parents_input(parents_df)

    parents_df["parent_id"] = parents_df["parent_id"].apply(clean_value)
    gardens_df["garden_id"] = gardens_df["garden_id"].apply(clean_value)

    capacities = {
        clean_value(row.get("garden_id", "")): to_int(
            row.get("capacity", DEFAULT_CAPACITY),
            default=DEFAULT_CAPACITY
        )
        for _, row in gardens_df.iterrows()
        if clean_value(row.get("garden_id", ""))
    }

    (
        parent_preferences,
        garden_rankings,
        score_lookup,
        garden_priority_lookup,
        parent_preferences_df,
        garden_rankings_df,
        candidate_filtering_df,
        social_bonus_df
    ) = build_preferences(parents_df, gardens_df)

    parent_match, garden_matches, proposal_log_df = gale_shapley_with_capacities(
        parent_preferences,
        garden_rankings,
        capacities
    )

    assignment_source = {
        parent_id: "gale_shapley"
        for parent_id, garden_id in parent_match.items()
        if garden_id is not None
    }

    mandatory_assignment_df, planning_review_df = mandatory_nearby_assignment(
        parents_df,
        gardens_df,
        parent_match,
        garden_matches,
        capacities,
        score_lookup,
        garden_priority_lookup
    )

    for parent_id, garden_id in parent_match.items():
        if garden_id is not None and parent_id not in assignment_source:
            assignment_source[parent_id] = "mandatory_nearby_assignment"

    for _, row in planning_review_df.iterrows():
        assignment_source[clean_value(row.get("parent_id", ""))] = "requires_planning_review"

    # בדיקת יציבות GS טהור - לפני שלב ה-manual exception
    stability_summary_df, blocking_pairs_df = build_stability_check(
        parent_preferences,
        garden_rankings,
        parent_match,
        garden_matches,
        capacities,
        parents_df,
        gardens_df
    )

    manual_exception_log_df = apply_manual_exception_after_all_close_solution(
        parents_df,
        gardens_df,
        parent_match,
        garden_matches,
        capacities,
        score_lookup,
        garden_priority_lookup,
        planning_review_df
    )

    if not manual_exception_log_df.empty and "decision" in manual_exception_log_df.columns:
        for _, row in manual_exception_log_df.iterrows():
            if row.get("decision") == "applied":
                assignment_source[clean_value(row.get("parent_id", ""))] = "final_manual_exception"

    distance_audit_df = build_distance_audit(
        parents_df,
        gardens_df,
        parent_match,
        assignment_source
    )

    create_outputs(
        parents_df,
        gardens_df,
        parent_match,
        garden_matches,
        score_lookup,
        garden_priority_lookup,
        parent_preferences_df,
        garden_rankings_df,
        candidate_filtering_df,
        proposal_log_df,
        stability_summary_df,
        blocking_pairs_df,
        capacities,
        mandatory_assignment_df,
        planning_review_df,
        manual_exception_log_df,
        distance_audit_df,
        assignment_source,
        social_bonus_df
    )


if __name__ == "__main__":
    main()
