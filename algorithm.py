import pandas as pd
from thefuzz import process

# 1. קריאת בסיס הנתונים (מקור האמת)
# נשתמש בגיליון המרכזי של הגנים. אם הקובץ המקורי שלך הוא אקסל, אפשר להשתמש ב- pd.read_excel('Database.xlsx')
try:
    df = pd.read_csv("Database.xlsx - Kindergartens.csv")
    
    # נוודא שאנחנו עובדים רק עם שורות שיש בהן שם של גן (מנקים שורות ריקות)
    # הערה: נצטרך לעדכן את השם 'שם_הגן' לשם העמודה המדויק שיש לכן בקובץ!
    df = df.dropna(subset=['שם_הגן'])
    
    # יצירת רשימה של כל שמות הגנים מתוך ה-Database
    db_kindergartens = df['שם_הגן'].tolist()

    # 2. קלט מהמשתמש (חיפוש עם שגיאות או חוסרים)
    user_input = "גן שולה תל אביב"
    print(f"החיפוש שהוזן: '{user_input}'\n")

    # 3. הפעלת מנוע החיפוש וההתאמה
    best_match = process.extractOne(user_input, db_kindergartens)
    matched_name = best_match[0]
    score = best_match[1]

    # 4. לוגיקת החלטה (Threshold) - המרכיב החשוב בחקר ביצועים
    THRESHOLD = 80  # ציון המינימום הנדרש כדי לאשר שזו אותה ישות

    if score >= THRESHOLD:
        print("✅ נמצאה התאמה במאגר!")
        print(f"שם הגן במקור האמת: {matched_name} (ציון: {score}/100)")
        
        # בונוס: שליפת כל שאר הפרטים (כתובת, טלפון וכו') של הגן שנמצא
        kindergarten_info = df[df['שם_הגן'] == matched_name]
        print("\nפרטי הגן המלאים:")
        print(kindergarten_info)
        
    else:
        print("❌ לא נמצאה התאמה ודאית.")
        print(f"הגן הכי קרוב שמצאנו הוא '{matched_name}' אבל הציון נמוך מדי ({score}/100).")

except FileNotFoundError:
    print("שגיאה: לא הצלחתי למצוא את קובץ הנתונים בתיקייה.")
except KeyError:
    print("שגיאה: לא מצאתי עמודה בשם הזה. צריך לעדכן את שם העמודה בקוד.")