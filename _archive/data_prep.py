import pandas as pd
import re

# ייבוא כל פונקציות העזר וההשוואה המקוריות שלך מתוך קובץ data_utils
from data_utils import (
    normalize_text, is_empty, completeness_score, merge_values, 
    merge_license_status, name_match, rows_are_duplicates, extract_phone)

try:
    # טעינת נתונים ובדיקת תקינות
    # Loading data and checking validity
    df = pd.read_excel("Database.xlsx", sheet_name=0)

    expected_source_columns = ['שם וסמל מעון', 'סטטוס הרישוי', 'טלפון',
                               'בעלות', 'מגזר', 'ישוב', 'כתובת', 'מנהל/ת המעון']

    missing_cols = [col for col in expected_source_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    # סינון ראשוני ופיצול עמודות
    # Initial filtering and column splitting
    
    # סינון גנים שנסגרו
    # Filtering closed kindergartens
    df = df[df['סטטוס הרישוי'] != 'מעון סגור'].copy()

    # חילוץ סמל המעון מתוך השם אם קיים
    # Extracting the daycare ID from the name if exists
    df['id'] = df['שם וסמל מעון'].astype(str).str.extract(r'(\d+)')

    # ניקוי השם ממספרים ותווים מיותרים
    # Cleaning the name from numbers and unnecessary characters
    df['name'] = (
        df['שם וסמל מעון']
        .astype(str)
        .str.replace(r'\d+', '', regex=True)
        .str.replace('-', '', regex=False)
        .str.strip()
    )

    # ארגון ושינוי שמות עמודות
    # Organizing and renaming columns
    rename_map = {
        'בעלות': 'ownership',
        'מגזר': 'sector',
        'סטטוס הרישוי': 'license_status',
        'ישוב': 'city',
        'כתובת': 'address',
        'מנהל/ת המעון': 'manager'
    }
    df = df.rename(columns=rename_map)
    df = df.drop(columns=[col for col in ['שם וסמל מעון', 'סמל זרוע העבודה'] if col in df.columns])

    # ניקוי נתונים בסיסי
    # Basic data cleaning
    text_cols = ['name', 'city', 'address', 'ownership', 'sector', 'manager', 'license_status']

    # מילוי ערכים ריקים בלא ידוע
    # Filling empty values with Unknown
    for col in text_cols:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace(['nan', 'None', ''], 'Unknown')

    # ניקוי מספרי טלפון באמצעות פונקציית העזר המשותפת
    # Cleaning phone numbers using the shared helper function
    df['phone'] = df['טלפון'].apply(extract_phone)
    df['phone'] = df['phone'].replace('', 'Unknown')
    
    if 'טלפון' in df.columns:
        df = df.drop(columns=['טלפון'])

    # פונקציות לוגיות ספציפיות לתהליך האיחוד (הן נשארות כאן כי אינן גנריות)
    # Logical functions specific to the merging process (they remain here as they are not generic)

    # הכנת עמודות עזר לאיחוד תוך שימוש בפונקציות המיובאות
    # Preparing helper columns for merging using imported functions
    df['name_norm'] = df['name'].apply(lambda x: normalize_text(x, remove_numbers=True))
    df['city_norm'] = df['city'].apply(lambda x: normalize_text(x, remove_numbers=False))
    df['address_norm'] = df['address'].apply(lambda x: normalize_text(x, remove_numbers=False))
    df['manager_norm'] = df['manager'].apply(lambda x: normalize_text(x, remove_numbers=True))
    df['ownership_norm'] = df['ownership'].apply(lambda x: normalize_text(x, remove_numbers=True))
    df['has_real_id'] = df['id'].notna() & (df['id'] != 'Unknown') & (df['id'] != '')
    df['completeness_score'] = df.apply(completeness_score, axis=1)

    # סידור הנתונים כך שהרשומות המלאות ביותר והרשמיות ביותר יופיעו ראשונות
    # Sorting the data so the most complete and official records appear first
    df = df.sort_values(by=['has_real_id', 'completeness_score'], ascending=[False, False])

    # מנוע איחוד הישויות
    # Entity resolution engine
    merged_rows = []
    recycle_bin_rows = []

    # רץ על כל קבוצת כתובות זהה בניסיון למצוא כפילויות
    # Iterating over each identical address group in an attempt to find duplicates
    for _, address_group in df.groupby(['city_norm', 'address_norm'], dropna=False):
        address_group = address_group.copy()
        used_indexes = set()

        for idx, row in address_group.iterrows():
            if idx in used_indexes: continue

            cluster_indexes = [idx]
            used_indexes.add(idx)
            changed = True

            # בניית אשכול של רשומות מקושרות
            # Building a cluster of linked records
            while changed:
                changed = False
                for other_idx, other_row in address_group.iterrows():
                    if other_idx in used_indexes: continue

                    for cluster_idx in cluster_indexes:
                        cluster_row = address_group.loc[cluster_idx]
                        if rows_are_duplicates(cluster_row, other_row):
                            cluster_indexes.append(other_idx)
                            used_indexes.add(other_idx)
                            changed = True
                            break

            similar_group = address_group.loc[cluster_indexes].copy()
            similar_group = similar_group.sort_values(
                by=['has_real_id', 'completeness_score'], ascending=[False, False])

            # פיצול לאיחוד, השורה המלאה ביותר קולטת אליה את הנתונים והשאר נשלחות למחזור
            # Splitting for merge, the most complete row absorbs the data and the rest are sent to recycling
            main_row = similar_group.iloc[0].copy()
            duplicate_rows = similar_group.iloc[1:].copy()

            for col in df.columns:
                if col in ['id', 'city', 'address', 'name_norm', 'city_norm', 'address_norm',
                           'manager_norm', 'ownership_norm', 'has_real_id', 'completeness_score']:
                    continue

                if col == 'license_status':
                    main_row[col] = merge_license_status(similar_group[col])
                else:
                    main_row[col] = merge_values(similar_group[col])

            merged_rows.append(main_row)

            if not duplicate_rows.empty:
                duplicate_rows['duplicate_reason'] = 'Merged due to matching location and one strong indicator'
                recycle_bin_rows.append(duplicate_rows)

    # סיכום ויצוא קבצים
    # Summary and file export
    df = pd.DataFrame(merged_rows)

    if recycle_bin_rows:
        recycle_bin_df = pd.concat(recycle_bin_rows, ignore_index=True)
    else:
        recycle_bin_df = pd.DataFrame(columns=list(df.columns) + ['duplicate_reason'])

    # ניקוי עמודות עזר ששימשו לאיחוד בלבד
    # Cleaning helper columns used for merging only
    helper_cols = ['name_norm', 'city_norm', 'address_norm', 'manager_norm', 
                   'ownership_norm', 'has_real_id', 'completeness_score']
    df = df.drop(columns=helper_cols, errors='ignore')
    recycle_bin_df = recycle_bin_df.drop(columns=helper_cols, errors='ignore')

    # סידור עמודות סופי
    # Final column sorting
    cols_order = ['id', 'name', 'city', 'address', 'phone', 
                  'ownership', 'sector', 'manager', 'license_status']
    
    # הבטחה שכל העמודות הנדרשות קיימות בטבלה והסדר שלהן מדויק
    # Ensuring all required columns exist in the table and their order is exact
    df = df[[col for col in cols_order if col in df.columns]]
    recycle_cols_order = cols_order + ['duplicate_reason']
    recycle_bin_df = recycle_bin_df[[col for col in recycle_cols_order if col in recycle_bin_df.columns]]

    # שמירה לאקסל
    # Saving to Excel
    with pd.ExcelWriter("Cleaned_Database.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        recycle_bin_df.to_excel(writer, sheet_name="Duplicates_Recycle_Bin", index=False)

    print(f"Done | Active: {len(df)} | Recycle: {len(recycle_bin_df)}")

except FileNotFoundError:
    print("Database.xlsx not found.")
except Exception as e:
    print(f"Error occurred: {e}")