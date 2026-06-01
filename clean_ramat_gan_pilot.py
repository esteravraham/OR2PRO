import pandas as pd
import random
import re
from difflib import SequenceMatcher

try:
    df = pd.read_excel("Ramat_Gan_Pilot.xlsx", sheet_name="Official_Data")

    expected_source_columns = [
        'שם וסמל מעון',
        'סטטוס הרישוי',
        'טלפון',
        'בעלות',
        'מגזר',
        'ישוב',
        'כתובת',
        'מנהל/ת המעון'
    ]

    missing_cols = [col for col in expected_source_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    df = df[df['סטטוס הרישוי'] != 'מעון סגור'].copy()

    df['id'] = df['שם וסמל מעון'].astype(str).str.extract(r'(\d+)')

    df['name'] = (
        df['שם וסמל מעון']
        .astype(str)
        .str.replace(r'\d+', '', regex=True)
        .str.replace('-', '', regex=False)
        .str.strip()
    )

    rename_map = {
        'בעלות': 'ownership',
        'מגזר': 'sector',
        'סטטוס הרישוי': 'license_status',
        'ישוב': 'city',
        'כתובת': 'address',
        'מנהל/ת המעון': 'manager'
    }

    df = df.rename(columns=rename_map)

    df = df.drop(
        columns=[col for col in ['שם וסמל מעון', 'סמל זרוע העבודה'] if col in df.columns]
    )

    text_cols = ['name', 'city', 'address', 'ownership', 'sector', 'manager', 'license_status']

    for col in text_cols:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(['nan', 'None', ''], 'Unknown')

    df['phone'] = df['טלפון'].astype(str).str.replace(r'[^0-9]', '', regex=True)
    df['phone'] = df['phone'].replace(['nan', 'None', ''], None)

    df['phone'] = df['phone'].apply(
        lambda x: f"0{x[3:]}" if pd.notna(x) and str(x).startswith('972') else x
    )

    df['phone'] = df['phone'].apply(
        lambda x: f"0{x}" if pd.notna(x) and not str(x).startswith('0') else x
    )

    df['phone'] = df['phone'].fillna('Unknown')

    if 'טלפון' in df.columns:
        df = df.drop(columns=['טלפון'])

    def normalize_text(value, remove_numbers=True):
        if pd.isna(value):
            return ''

        value = str(value).strip()
        if remove_numbers:
            value = re.sub(r'\d+', '', value)
        value = re.sub(r'[-–—״"\'.,()]', ' ', value)
        value = re.sub(r'\s+', ' ', value)

        return value.strip()

    def is_empty(value):
        return pd.isna(value) or value == '' or value == 'Unknown'

    def completeness_score(row):
        return sum(1 for value in row if not is_empty(value))

    def merge_values(values):
        clean_values = []

        for value in values:
            if not is_empty(value):
                value = str(value).strip()
                if value not in clean_values:
                    clean_values.append(value)

        if not clean_values:
            return 'Unknown'

        return ' | '.join(clean_values)

    def merge_license_status(values):
        status_priority = {
            'רישיון בתוקף': 1,
            'בתהליך רישוי': 2,
            'לא הוגשה בקשה לרישוי': 3,
            'Unknown': 9
        }

        clean_values = [value for value in values if not is_empty(value)]

        if not clean_values:
            return 'Unknown'

        return min(clean_values, key=lambda value: status_priority.get(value, 9))

    def name_match(name1, name2):
        if is_empty(name1) or is_empty(name2):
            return False

        name1 = normalize_text(name1)
        name2 = normalize_text(name2)

        if name1 == '' or name2 == '':
            return False

        if name1 == name2:
            return True

        short_name = min(name1, name2, key=len)
        long_name = max(name1, name2, key=len)

        if len(short_name) < 4:
            return False

        return short_name in long_name

    def exact_strong_match(value1, value2):
        if is_empty(value1) or is_empty(value2):
            return False

        value1 = normalize_text(value1)
        value2 = normalize_text(value2)

        if value1 == '' or value2 == '':
            return False

        return value1 == value2

    def ownership_match(value1, value2):
        if is_empty(value1) or is_empty(value2):
            return False

        value1 = normalize_text(value1)
        value2 = normalize_text(value2)

        weak_values = {'פרטי', 'ציבורי', 'עירוני', 'עמותה', 'חברה', 'Unknown'}

        if value1 in weak_values or value2 in weak_values:
            return False

        return value1 == value2

    def rows_are_duplicates(row1, row2):
        if row1['city_norm'] != row2['city_norm']:
            return False

        if row1['address_norm'] != row2['address_norm']:
            return False

        if exact_strong_match(row1['phone'], row2['phone']):
            return True

        if name_match(row1['name_norm'], row2['name_norm']):
            return True

        if exact_strong_match(row1['manager_norm'], row2['manager_norm']):
            return True

        if ownership_match(row1['ownership_norm'], row2['ownership_norm']):
            return True

        return False

    df['name_norm'] = df['name'].apply(lambda x: normalize_text(x, remove_numbers=True))
    df['city_norm'] = df['city'].apply(lambda x: normalize_text(x, remove_numbers=False))
    df['address_norm'] = df['address'].apply(lambda x: normalize_text(x, remove_numbers=False))
    df['manager_norm'] = df['manager'].apply(lambda x: normalize_text(x, remove_numbers=True))
    df['ownership_norm'] = df['ownership'].apply(lambda x: normalize_text(x, remove_numbers=True))
    df['has_real_id'] = df['id'].notna()
    df['completeness_score'] = df.apply(completeness_score, axis=1)

    df = df.sort_values(
        by=['has_real_id', 'completeness_score'],
        ascending=[False, False]
    )

    merged_rows = []
    recycle_bin_rows = []

    for _, address_group in df.groupby(['city_norm', 'address_norm'], dropna=False):
        address_group = address_group.copy()
        used_indexes = set()

        for idx, row in address_group.iterrows():
            if idx in used_indexes:
                continue

            cluster_indexes = [idx]
            used_indexes.add(idx)

            changed = True

            while changed:
                changed = False

                for other_idx, other_row in address_group.iterrows():
                    if other_idx in used_indexes:
                        continue

                    for cluster_idx in cluster_indexes:
                        cluster_row = address_group.loc[cluster_idx]

                        if rows_are_duplicates(cluster_row, other_row):
                            cluster_indexes.append(other_idx)
                            used_indexes.add(other_idx)
                            changed = True
                            break

            similar_group = address_group.loc[cluster_indexes].copy()

            similar_group = similar_group.sort_values(
                by=['has_real_id', 'completeness_score'],
                ascending=[False, False]
            )

            main_row = similar_group.iloc[0].copy()
            duplicate_rows = similar_group.iloc[1:].copy()

            for col in df.columns:
                if col in [
                    'id',
                    'city',
                    'address',
                    'name_norm',
                    'city_norm',
                    'address_norm',
                    'manager_norm',
                    'ownership_norm',
                    'has_real_id',
                    'completeness_score'
                ]:
                    continue

                if col == 'license_status':
                    main_row[col] = merge_license_status(similar_group[col])
                else:
                    main_row[col] = merge_values(similar_group[col])

            merged_rows.append(main_row)

            if not duplicate_rows.empty:
                duplicate_rows['duplicate_reason'] = 'Same city and address with one strong matching field'
                recycle_bin_rows.append(duplicate_rows)

    df = pd.DataFrame(merged_rows)

    if recycle_bin_rows:
        recycle_bin_df = pd.concat(recycle_bin_rows, ignore_index=True)
    else:
        recycle_bin_df = pd.DataFrame(columns=list(df.columns) + ['duplicate_reason'])

    helper_cols = [
        'name_norm',
        'city_norm',
        'address_norm',
        'manager_norm',
        'ownership_norm',
        'has_real_id',
        'completeness_score'
    ]

    df = df.drop(columns=helper_cols, errors='ignore')
    recycle_bin_df = recycle_bin_df.drop(columns=helper_cols, errors='ignore')

    existing_ids = set(df['id'].dropna().astype(str).unique())

    def generate_unique_id():
        while True:
            new_id = str(random.randint(10000, 99999))
            if new_id not in existing_ids:
                existing_ids.add(new_id)
                return new_id

    df['id'] = df['id'].apply(
        lambda x: generate_unique_id() if pd.isna(x) or x == '' or x == 'Unknown' else str(x)
    )

    cols_order = [
        'id',
        'name',
        'city',
        'address',
        'phone',
        'ownership',
        'sector',
        'manager',
        'license_status'
    ]

    df = df[cols_order]

    recycle_cols_order = cols_order + ['duplicate_reason']
    recycle_bin_df = recycle_bin_df[
        [col for col in recycle_cols_order if col in recycle_bin_df.columns]
    ]

    def text_similarity(value1, value2):
        if pd.isna(value1) or pd.isna(value2):
            return 0

        value1 = normalize_text(value1, remove_numbers=True)
        value2 = normalize_text(value2, remove_numbers=True)

        if value1 == "" or value2 == "":
            return 0

        return SequenceMatcher(None, value1, value2).ratio()

    possible_official_duplicates = []

    for i in range(len(df)):
        for j in range(i + 1, len(df)):
            row1 = df.iloc[i]
            row2 = df.iloc[j]

            if row1["city"] != row2["city"]:
                continue

            same_phone = row1["phone"] == row2["phone"] and row1["phone"] != "Unknown"
            same_ownership = row1["ownership"] == row2["ownership"] and row1["ownership"] != "Unknown"
            name_sim = text_similarity(row1["name"], row2["name"])

            if name_sim >= 0.75 and (same_phone or same_ownership):
                possible_official_duplicates.append({
                    "id_1": row1["id"],
                    "name_1": row1["name"],
                    "address_1": row1["address"],
                    "phone_1": row1["phone"],
                    "ownership_1": row1["ownership"],
                    "manager_1": row1["manager"],

                    "id_2": row2["id"],
                    "name_2": row2["name"],
                    "address_2": row2["address"],
                    "phone_2": row2["phone"],
                    "ownership_2": row2["ownership"],
                    "manager_2": row2["manager"],

                    "name_similarity": round(name_sim, 2),
                    "same_phone": same_phone,
                    "same_ownership": same_ownership,
                    "review_decision": ""
                })

    possible_official_duplicates_df = pd.DataFrame(possible_official_duplicates)

    with pd.ExcelWriter("Clean_Ramat_Gan_Pilot.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        recycle_bin_df.to_excel(writer, sheet_name="Duplicates_Recycle_Bin", index=False)
        possible_official_duplicates_df.to_excel(
            writer,
            sheet_name="Official_Possible_Duplicates",
            index=False
        )

    print(
        f"✅ Done | Active: {len(df)} | "
        f"Recycle: {len(recycle_bin_df)} | "
        f"Possible official duplicates: {len(possible_official_duplicates_df)}"
    )

except FileNotFoundError:
    print("❌ Ramat_Gan_Pilot.xlsx not found")

except Exception as e:
    print(f"❌ Error: {e}")