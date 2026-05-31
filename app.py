from flask import Flask, request, render_template_string, send_from_directory
import pandas as pd

app = Flask(__name__)

EXCEL_FILE = "Cleaned_Database.xlsx"

def load_data():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Cleaned_Data")
    except ValueError:
        df = pd.read_excel(EXCEL_FILE, sheet_name=0)

    df = df.fillna("Unknown")
    df["id"] = df["id"].astype(str)
    return df

@app.route("/")
def index():
    return send_from_directory(".", "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(".", filename)

@app.route("/search")
def search():
    query = request.args.get("q", "").strip()
    df = load_data()

    if query:
        results = df[
            df["name"].astype(str).str.contains(query, case=False, na=False) |
            df["city"].astype(str).str.contains(query, case=False, na=False) |
            df["address"].astype(str).str.contains(query, case=False, na=False)
        ]
    else:
        results = df

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>תוצאות חיפוש</title>
        <link rel="stylesheet" href="/optionA.css">
    </head>
    <body class="inner-page">
        <h1>תוצאות חיפוש</h1>

        {% if query %}
            <p>נמצאו {{ results|length }} תוצאות עבור: <strong>{{ query }}</strong></p>
        {% else %}
            <p>מציג את כל הגנים במאגר</p>
        {% endif %}

        <table class="data-table">
            <thead>
                <tr>
                    <th>שם הגן</th>
                    <th>עיר</th>
                    <th>כתובת</th>
                    <th>סטטוס רישוי</th>
                    <th>פרופיל</th>
                </tr>
            </thead>
            <tbody>
                {% for garden in results %}
                <tr>
                    <td>{{ garden.name }}</td>
                    <td>{{ garden.city }}</td>
                    <td>{{ garden.address }}</td>
                    <td>{{ garden.license_status }}</td>
                    <td><a href="/garden/{{ garden.id }}">צפייה</a></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </body>
    </html>
    """, results=results.to_dict(orient="records"), query=query)

@app.route("/garden/<garden_id>")
def garden_profile(garden_id):
    df = load_data()
    garden = df[df["id"] == str(garden_id)]

    if garden.empty:
        return "גן לא נמצא"

    garden = garden.iloc[0].to_dict()

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>פרופיל גן</title>
        <link rel="stylesheet" href="/optionA.css">
    </head>
    <body class="inner-page">

        <section class="garden-profile">

            <div class="profile-header">
                <div>
                    <h1>{{ garden.name }}</h1>
                    <p class="profile-subtitle">{{ garden.city }} · {{ garden.address }}</p>
                </div>
                <span class="license-badge">{{ garden.license_status }}</span>
            </div>

            <div class="profile-gallery">
                <div class="gallery-placeholder">תמונה 1</div>
                <div class="gallery-placeholder">תמונה 2</div>
                <div class="gallery-placeholder">תמונה 3</div>
            </div>

            <div class="profile-grid">

                <section class="profile-card">
                    <h2>מידע רשמי</h2>
                    <p><strong>טלפון:</strong> {{ garden.phone }}</p>
                    <p><strong>בעלות:</strong> {{ garden.ownership }}</p>
                    <p><strong>מגזר:</strong> {{ garden.sector }}</p>
                    <p><strong>מנהלת:</strong> {{ garden.manager }}</p>
                </section>

                <section class="profile-card">
                    <h2>מידע משלים</h2>
                    <p><strong>שעות פעילות:</strong> לא ידוע</p>
                    <p><strong>אתר:</strong> לא נמצא</p>
                    <p><strong>שכונה:</strong> לא ידוע</p>
                    <p><strong>מקורות:</strong> משרד החינוך</p>
                </section>

                <section class="profile-card">
                    <h2>מידע מהקהילה</h2>
                    <p><strong>מצלמות להורים:</strong> לא ידוע</p>
                    <p><strong>מרחב מוגן:</strong> לא ידוע</p>
                    <p><strong>מחיר חודשי:</strong> לא ידוע</p>
                    <p><strong>יחס צוות-ילדים:</strong> לא ידוע</p>
                </section>

                <section class="profile-card">
                    <h2>חוות דעת הורים</h2>
                    <p>עדיין לא נוספו חוות דעת.</p>
                    <button class="secondary-button">הוספת חוות דעת</button>
                </section>

            </div>

        </section>

    </body>
    </html>
    """, garden=garden)

if __name__ == "__main__":
    app.run(debug=True)