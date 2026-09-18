#!/usr/bin/env python3
"""Small local website: one button fetches the latest news, the page shows it sorted by keywords.

Run:  python3 app.py   then open http://127.0.0.1:5000
"""

import os
import subprocess
import sys

from flask import Flask, jsonify, render_template

import classify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FETCH_SCRIPT = os.path.join(BASE_DIR, "news_rss.py")
XLSX_PATH = os.path.join(BASE_DIR, "news_articles.xlsx")

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


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
