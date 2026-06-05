from flask import Flask, request, jsonify, send_from_directory
import json
import os
from datetime import datetime

app = Flask(__name__, static_folder=".", static_url_path="")

SUBMISSIONS_FILE = "parent_reviews_submissions.json"
GARDENS_FILE = "gardens.json"


def load_json_file(file_path, default_value):
    if not os.path.exists(file_path):
        return default_value

    with open(file_path, "r", encoding="utf-8") as file:
        return json.load(file)


def save_json_file(file_path, data):
    with open(file_path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def garden_exists(garden_id):
    data = load_json_file(GARDENS_FILE, {"gardens": []})
    gardens = data.get("gardens", [])

    return any(str(garden.get("id")) == str(garden_id) for garden in gardens)


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/<path:filename>")
def serve_static_file(filename):
    return send_from_directory(".", filename)


@app.route("/submit_review", methods=["POST"])
def submit_review():
    review_data = request.get_json()

    if not review_data:
        return jsonify({
            "success": False,
            "message": "No review data received"
        }), 400

    garden_id = review_data.get("garden_id")

    if not garden_id:
        return jsonify({
            "success": False,
            "message": "Missing garden_id"
        }), 400

    if not garden_exists(garden_id):
        return jsonify({
            "success": False,
            "message": "Garden ID not found"
        }), 404

    submissions = load_json_file(SUBMISSIONS_FILE, [])

    review_data["submission_id"] = len(submissions) + 1
    review_data["server_created_at"] = datetime.now().isoformat(timespec="seconds")
    review_data["status"] = "pending_review"
    review_data["source"] = "website_form"

    submissions.append(review_data)

    save_json_file(SUBMISSIONS_FILE, submissions)

    return jsonify({
        "success": True,
        "message": "Review submitted successfully",
        "submission_id": review_data["submission_id"]
    })


@app.route("/admin/submissions")
def view_submissions():
    submissions = load_json_file(SUBMISSIONS_FILE, [])

    return jsonify({
        "total_submissions": len(submissions),
        "submissions": submissions
    })


if __name__ == "__main__":
    app.run(debug=True)