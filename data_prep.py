import pandas as pd
import random

try:
    df = pd.read_excel("Database.xlsx", sheet_name=0)
    original_row_count = len(df)

    expected_source_columns = ['שם וסמל מעון', 'סטטוס הרישוי', 'טלפון', 'בעלות', 'מגזר', 'ישוב', 'כתובת', 'מנהל/ת המעון']
    
    missing_cols = [col for col in expected_source_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"ERROR: The following required columns are missing from the source file: {missing_cols}.")

    
    df = df[df['סטטוס הרישוי'] != 'מעון סגור']

    df['id'] = df['שם וסמל מעון'].astype(str).str.extract(r'(\d+)')
    
    existing_ids = set(df['id'].dropna().unique())
    def generate_unique_id():
        while True:
            new_id = str(random.randint(10000, 99999))
            if new_id not in existing_ids:
                existing_ids.add(new_id)
                return new_id
    
    df['id'] = df['id'].apply(lambda x: generate_unique_id() if pd.isna(x) else x)
    
    df['name'] = df['שם וסמל מעון'].astype(str).str.replace(r'\d+', '', regex=True).str.replace('-', '', regex=False).str.strip()
    
    if 'שם וסמל מעון' in df.columns:
        df = df.drop(columns=['שם וסמל מעון'])
    if 'סמל זרוע העבודה' in df.columns:
        df = df.drop(columns=['סמל זרוע העבודה'])

    df['phone'] = df['טלפון'].astype(str).str.replace(r'[^0-9]', '', regex=True)
    df['phone'] = df['phone'].replace(['nan', 'None', ''], None)
    df['phone'] = df['phone'].apply(lambda x: f"0{x[3:]}" if pd.notna(x) and str(x).startswith('972') else x)
    df['phone'] = df['phone'].apply(lambda x: f"0{x}" if pd.notna(x) and not str(x).startswith('0') else x)
    df['phone'] = df['phone'].fillna('Unknown')

    rename_map = {
        'בעלות': 'ownership',
        'מגזר': 'sector',
        'סטטוס הרישוי': 'license_status',
        'ישוב': 'city',
        'כתובת': 'address',
        'מנהל/ת המעון': 'manager'
    }
    df = df.rename(columns=rename_map)

    df['ownership'] = df['ownership'].fillna('Unknown')
    df['sector'] = df['sector'].fillna('Unknown')
    df['manager'] = df['manager'].fillna('Unknown')

    cols_order = ['id', 'name', 'city', 'address', 'phone', 'ownership', 'sector', 'manager', 'license_status']
    
    missing_final_cols = [col for col in cols_order if col not in df.columns]
    if missing_final_cols:
        raise ValueError(f"ERROR: Processing failed to create these columns: {missing_final_cols}")

    df = df[cols_order]

    df.to_excel("Cleaned_Database.xlsx", index=False)
    

    print("\n✅ Success: Data processed and validated successfully.")
    

except FileNotFoundError:
    print("\n❌ Error: Could not find Database.xlsx.\n")
except Exception as e:
    print(f"\n❌ SCRIPT HALTED - ERROR DETECTED:\n{e}\n")