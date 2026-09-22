#!/usr/bin/env python3
"""Small local website: one button fetches the latest news, the page shows it sorted by keywords.

Run:  python3 app.py   then open http://127.0.0.1:5000
"""

import json
import os
import subprocess
import sys
import threading
from datetime import datetime

from flask import Flask, jsonify, render_template, request

import classify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FETCH_SCRIPT = os.path.join(BASE_DIR, "news_rss.py")
XLSX_PATH = os.path.join(BASE_DIR, "news_articles.xlsx")
# Articles the user marked as read, with their rating. Keyed by article link because the
# numeric ids from classify.py change on every fetch and the xlsx is overwritten each time.
HISTORY_PATH = os.path.join(BASE_DIR, "reading_history.json")

RATINGS = ["great", "normal", "hard", "bad"]
# Flask serves requests in threads; two quick clicks would otherwise both load the file, each add
# their own entry and the second save would overwrite the first. So every change holds this lock.
HISTORY_LOCK = threading.Lock()

# Fetching all feeds usually takes 1-2 minutes
FETCH_TIMEOUT_SECONDS = 300

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/refresh", methods=["POST"])
def refresh():
    """Run news_rss.py to download the latest articles into the Excel file."""
    try:
        result = subprocess.run(
            [sys.executable, FETCH_SCRIPT],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=FETCH_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "log": "Fetching took too long and was stopped."})

    log = (result.stdout + result.stderr).strip().splitlines()
    return jsonify({"ok": result.returncode == 0, "log": "\n".join(log[-15:])})


@app.route("/articles")
def articles():
    """Return all articles, already classified, as JSON."""
    if not os.path.exists(XLSX_PATH):
        return jsonify({"ok": False, "message": "No news yet. Click \"Check latest news\" first."})
    return jsonify(classify.build_report(XLSX_PATH))


# ---- Reading history --------------------------------------------------------

def load_history():
    """Read reading_history.json; a missing or broken file just means an empty history."""
    try:
        with open(HISTORY_PATH, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("articles"), dict):
            return data
    except (OSError, ValueError):
        pass
    return {"articles": {}}


def save_history(data):
    """Write to a temporary file first, then rename, so a crash mid-write can't leave a half file."""
    temp_path = HISTORY_PATH + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(temp_path, HISTORY_PATH)


@app.route("/history")
def history():
    return jsonify(load_history())


@app.route("/history/add", methods=["POST"])
def history_add():
    """Mark an article as read. The body is a snapshot of the article so it can still be shown later."""
    body = request.get_json(silent=True) or {}
    link = body.get("link")
    if not link:
        return jsonify({"ok": False, "message": "Missing link"}), 400

    with HISTORY_LOCK:
        data = load_history()
        if link in data["articles"]:
            # Already read: keep the stored snapshot, rating and original read time
            return jsonify({"ok": True, "entry": data["articles"][link]})
        entry = {
            "link": link,
            "title": body.get("title", ""),
            "description": body.get("description", ""),
            "publication": body.get("publication", ""),
            "section": body.get("section", ""),
            "author": body.get("author", ""),
            "published": body.get("published", ""),
            "topic": body.get("topic", ""),
            # Same format as "published" so the page can sort them the same way
            "read_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "rating": None,
        }
        data["articles"][link] = entry
        save_history(data)
    return jsonify({"ok": True, "entry": entry})


@app.route("/history/rate", methods=["POST"])
def history_rate():
    body = request.get_json(silent=True) or {}
    link = body.get("link")
    rating = body.get("rating")
    if rating is not None and rating not in RATINGS:
        return jsonify({"ok": False, "message": "Unknown rating"}), 400

    with HISTORY_LOCK:
        data = load_history()
        if link not in data["articles"]:
            return jsonify({"ok": False, "message": "Article is not in the history"}), 404
        data["articles"][link]["rating"] = rating
        save_history(data)
    return jsonify({"ok": True, "entry": data["articles"][link]})


@app.route("/history/remove", methods=["POST"])
def history_remove():
    body = request.get_json(silent=True) or {}
    link = body.get("link")
    with HISTORY_LOCK:
        data = load_history()
        data["articles"].pop(link, None)
        save_history(data)
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
