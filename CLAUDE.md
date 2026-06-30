# OR2PRO — מערכת מידע לגני ילדים

## קבצי ליבה
- `app.py` — שרת Flask ראשי (382 שורות), כל ה-routes
- `algorithm.py` — אלגוריתם Gale-Shapley להתאמת הורים לגנים
- `data_utils.py` — פונקציות עזר, נייבא ע"י app.py
- `gardens.json` — בסיס נתונים ראשי של הגנים
- `parent_reviews_submissions.json` — ביקורות הורים, נכתב בזמן ריצה
- `admin_run_log.json` — לוג ריצות אלגוריתם, נכתב בזמן ריצה
- `Parents_Input.xlsx` — קלט הורים, נקרא/נכתב בזמן ריצה
- `Gale_Shapley_Matching_Result.xlsx` — תוצאות ההתאמה, נוצר אחרי ריצת האלגוריתם

## תיקיות
- `templates/` — כל ה-HTML (9 קבצים) + optionA.css
- `images/` — תמונות לאתר
- `_archive/` — סקריפטים ישנים וקבצי Excel שלא בשימוש שוטף (לא לגעת)

## טכנולוגיות
- Python + Flask
- static_folder="." (שורש הפרויקט)
- template_folder="templates"
