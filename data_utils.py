#סקריפט עם פונקציות עזר שימושיות


import re
import pandas as pd

# ----- הפונקציות המקוריות והמדויקות שלך מקובץ המקור -----

def normalize_text(value, remove_numbers=True):
    if pd.isna(value): return ''
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
    if not clean_values: return 'Unknown'
    return ' | '.join(clean_values)

def merge_license_status(values):
    status_priority = {
        'רישיון בתוקף': 1,
        'בתהליך רישוי': 2,
        'לא הוגשה בקשה לרישוי': 3,
        'Unknown': 9
    }
    clean_values = [value for value in values if not is_empty(value)]
    if not clean_values: return 'Unknown'
    return min(clean_values, key=lambda value: status_priority.get(value, 9))

def name_match(name1, name2):
    if is_empty(name1) or is_empty(name2): return False
    
    name1 = normalize_text(name1)
    name2 = normalize_text(name2)
    
    if name1 == '' or name2 == '': return False
    if name1 == name2: return True

    words1, words2 = set(name1.split()), set(name2.split())
    stop_words = {'גן', 'מעון', 'פעוטון', 'משפחתון', 'ילדים', 'הילדים', 'של', 'ה'}
    
    core_words1 = words1 - stop_words
    core_words2 = words2 - stop_words

    if not core_words1 or not core_words2: return False
    return core_words1.issubset(core_words2) or core_words2.issubset(core_words1)

def exact_strong_match(value1, value2):
    if is_empty(value1) or is_empty(value2): return False
    value1, value2 = normalize_text(value1), normalize_text(value2)
    if value1 == '' or value2 == '': return False
    return value1 == value2

def ownership_match(value1, value2):
    if is_empty(value1) or is_empty(value2): return False
    value1, value2 = normalize_text(value1), normalize_text(value2)
    weak_values = {'פרטי', 'ציבורי', 'עירוני', 'עמותה', 'חברה', 'Unknown'}
    if value1 in weak_values or value2 in weak_values: return False
    return value1 == value2

def rows_are_duplicates(row1, row2):
    if row1['city_norm'] != row2['city_norm']: return False
    if row1['address_norm'] != row2['address_norm']: return False

    if exact_strong_match(row1['phone'], row2['phone']): return True
    if name_match(row1['name_norm'], row2['name_norm']): return True
    if exact_strong_match(row1['manager_norm'], row2['manager_norm']): return True
    if ownership_match(row1['ownership_norm'], row2['ownership_norm']): return True
    
    return False

# ----- פונקציות העזר הכלליות שמשמשות את סקריפט ה-API -----

def extract_phone(value):
    if is_empty(value): return ''
    phone = re.sub(r'[^0-9]', '', str(value))
    if phone.startswith('972'): return f"0{phone[3:]}"
    if phone != '' and not phone.startswith('0'): return f"0{phone}"
    return phone

def address_match(addr1, addr2):
    if is_empty(addr1) or is_empty(addr2): return False
    
    # כאן אנו שולחים מפורשות remove_numbers=False כדי לשמור על מספרי הבתים ברחוב
    addr1 = normalize_text(addr1, remove_numbers=False)
    addr2 = normalize_text(addr2, remove_numbers=False)
    
    if addr1 == '' or addr2 == '': return False
    if addr1 == addr2: return True

    words1, words2 = set(addr1.split()), set(addr2.split())
    stop_words = {
        'רחוב', 'רח', 'שדרות', 'שד', 'דרך', 'סמטת', 'כיכר',
        'שכונת', 'שכונה', 'שכ', 'אוניברסיטת', 'אוניברסיטה', 'אונברסיטת', 'אונברסיטה', 'קמפוס'
    }
    core_words1 = words1 - stop_words
    core_words2 = words2 - stop_words

    if not core_words1 or not core_words2: return False
    return core_words1.issubset(core_words2) or core_words2.issubset(core_words1)


def api_db_match(api_address, db_address, api_phone, db_phone, api_name, db_name):
    if not address_match(api_address, db_address):
        return False

    if api_phone and db_phone and api_phone == db_phone:
        return True

    if name_match(api_name, db_name):
        return True

    return False